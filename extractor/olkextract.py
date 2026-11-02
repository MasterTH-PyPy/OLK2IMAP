#!/usr/bin/env python3
"""
olkextract - Extract mail from a "new Outlook" (Olk) IndexedDB offline store.

The new Outlook for Windows keeps NO classic mail store (no PST/OST). Its offline
data lives in the WebView2 IndexedDB of the OWA web app, under:
    %LocalAppData%\\Microsoft\\Olk\\EBWebView\\Default\\IndexedDB\\
        https_outlook.office.com_0.indexeddb.leveldb   (Chromium LevelDB)
        https_outlook.office.com_0.indexeddb.blob      (external blobs)
    %LocalAppData%\\Microsoft\\Olk\\Attachments\\       (cached attachment files)

This tool reads that store (Chromium LevelDB + V8 structured clone, via
ccl_chromium_reader), reconstructs the folder tree and messages, and writes:
    <out>/EML/<folder path>/<message>.eml      one RFC822 .eml per message
    <out>/manifest.json                        folder tree + per-message metadata

The C# front-end (OLK2IMAP) consumes manifest.json + the EML tree and pushes the
mail to IMAP / Exchange(EWS) / MBOX / PST, preserving folders and read state.

Only data that was cached OFFLINE can be recovered - not the full server mailbox.

Usage:
    olkextract --profile "C:\\Users\\<user>\\AppData\\Local\\Microsoft\\Olk" --out OUTDIR
  or point at the pieces explicitly:
    olkextract --leveldb DIR --blob DIR --attachments DIR --out OUTDIR

Dependency: ccl_chromium_reader
    pip install "git+https://github.com/cclgroupltd/ccl_chrome_indexeddb.git"
"""
import argparse, os, re, sys, json, shutil, tempfile, struct, glob, collections
import email.utils
from datetime import datetime
from email.message import EmailMessage

LDB_MAGIC = 0xdb4775248b80fb57
OFFLINE_DB_PREFIX = "owa-offline-data-"
ATTACH_DB_PREFIX = "owa-attachment-data-"


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def dec(b):
    """Olk stores many strings as raw one-byte (Windows-1252) buffers."""
    if isinstance(b, (bytes, bytearray)):
        return bytes(b).decode("cp1252", "replace")
    return b if isinstance(b, str) else ""


def _is_leveldb_dir(path):
    return os.path.isdir(path) and (
        os.path.isfile(os.path.join(path, "CURRENT")) or
        bool(glob.glob(os.path.join(path, "*.ldb"))) or
        bool(glob.glob(os.path.join(path, "*.log"))))


def _find_leveldb(base):
    """Find a *.indexeddb.leveldb directory anywhere under base. Prefer the
    outlook.office.com store; accept the standard Olk profile layout or a
    loose/extracted folder."""
    # fast path: standard Olk profile layout
    std = os.path.join(base, "EBWebView", "Default", "IndexedDB",
                       "https_outlook.office.com_0.indexeddb.leveldb")
    if _is_leveldb_dir(std):
        return std
    # the base itself may be a leveldb dir
    if _is_leveldb_dir(base):
        return base
    # otherwise search
    hits = []
    for root, dirs, _ in os.walk(base):
        for d in dirs:
            if d.endswith(".indexeddb.leveldb") and _is_leveldb_dir(os.path.join(root, d)):
                hits.append(os.path.join(root, d))
        if len(hits) > 50:
            break
    if not hits:
        return None
    hits.sort(key=lambda p: (0 if "outlook.office.com" in p else 1, len(p)))
    return hits[0]


def _find_sibling_blob(leveldb):
    if leveldb.endswith(".indexeddb.leveldb"):
        cand = leveldb[:-len(".indexeddb.leveldb")] + ".indexeddb.blob"
        if os.path.isdir(cand):
            return cand
    return None


def _find_attachments(base):
    direct = os.path.join(base, "Attachments")
    if os.path.isdir(direct):
        return direct
    for root, dirs, _ in os.walk(base):
        if "Attachments" in dirs:
            return os.path.join(root, "Attachments")
        # also accept a folder that already looks like the attachment store (ooa-*)
        for d in dirs:
            if d.startswith("ooa-"):
                return root
    return None


def resolve_paths(args):
    if args.leveldb:
        leveldb = args.leveldb
        blob = args.blob or _find_sibling_blob(leveldb)
        attachments = args.attachments
    elif args.profile:
        base = args.profile
        leveldb = _find_leveldb(base)
        if not leveldb:
            log(f"ERROR: no *.indexeddb.leveldb found under: {base}")
            log("  Point --profile at the Olk folder (…\\Microsoft\\Olk) or a backup that "
                "still contains EBWebView\\Default\\IndexedDB, or pass --leveldb directly.")
            sys.exit(2)
        blob = args.blob or _find_sibling_blob(leveldb)
        attachments = args.attachments or _find_attachments(base)
    else:
        log("ERROR: provide --profile or --leveldb.")
        sys.exit(2)

    if not leveldb or not os.path.isdir(leveldb):
        log(f"ERROR: leveldb directory not found: {leveldb}")
        sys.exit(2)
    if blob and not os.path.isdir(blob):
        blob = None
    if attachments and not os.path.isdir(attachments):
        attachments = None
    return leveldb, blob, attachments


def make_clean_leveldb(leveldb):
    """Copy the leveldb dir to a temp dir, dropping .ldb tables with a bad footer
    magic (truncated/locked during backup) so the reader doesn't abort."""
    tmp = tempfile.mkdtemp(prefix="olk_ldb_")
    bad = []
    for f in os.listdir(leveldb):
        src = os.path.join(leveldb, f)
        if not os.path.isfile(src):
            continue
        if f.endswith(".ldb"):
            try:
                with open(src, "rb") as fh:
                    fh.seek(-8, os.SEEK_END)
                    magic, = struct.unpack("<Q", fh.read(8))
                if magic != LDB_MAGIC:
                    bad.append(f)
                    continue
            except Exception:
                bad.append(f)
                continue
        shutil.copy2(src, os.path.join(tmp, f))
    if bad:
        log(f"  skipped {len(bad)} corrupt LevelDB table(s): {', '.join(bad)}")
    return tmp


def iter_store(db, name):
    """Iterate a store tolerantly; missing blobs / bad records are skipped."""
    os_ = db.get_object_store_by_name(name)
    stats = {"total": 0, "ok": 0, "bad": 0}

    def handler(key, rec):
        stats["bad"] += 1

    def gen():
        for rec in os_.iterate_records(live_only=False, bad_deserializer_data_handler=handler):
            stats["total"] += 1
            if rec.value is None:
                stats["bad"] += 1
                continue
            stats["ok"] += 1
            yield rec
    return gen(), stats


def safe(s):
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", s or "").strip().strip(".")
    return (s or "unnamed")[:80]


def addr(d):
    if not isinstance(d, dict):
        return None
    mb = d.get("Mailbox") if "Mailbox" in d else d
    if not isinstance(mb, dict):
        return None
    name = dec(mb.get("Name"))
    em = mb.get("EmailAddress") or ""
    em = dec(em) if isinstance(em, (bytes, bytearray)) else em
    return email.utils.formataddr((name, em)) if (em or name) else None


def addrs(lst):
    out = []
    if isinstance(lst, list):
        for x in lst:
            a = addr(x)
            if a:
                out.append(a)
    return ", ".join(out)


def main():
    ap = argparse.ArgumentParser(description="Extract mail from a new-Outlook (Olk) IndexedDB offline store.")
    ap.add_argument("--profile", help="Path to the Olk profile folder (auto-derives the rest)")
    ap.add_argument("--leveldb", help="Path to *.indexeddb.leveldb directory")
    ap.add_argument("--blob", help="Path to *.indexeddb.blob directory")
    ap.add_argument("--attachments", help="Path to the Olk Attachments directory")
    ap.add_argument("--out", required=True, help="Output directory")
    ap.add_argument("--json-progress", action="store_true", help="Emit JSON progress lines on stdout")
    args = ap.parse_args()

    try:
        from ccl_chromium_reader import ccl_chromium_indexeddb as idb
    except ImportError:
        log('ERROR: ccl_chromium_reader missing. Install with:\n'
            '  pip install "git+https://github.com/cclgroupltd/ccl_chrome_indexeddb.git"')
        sys.exit(3)

    leveldb, blob, attachments = resolve_paths(args)
    log(f"leveldb     : {leveldb}")
    log(f"blob        : {blob}")
    log(f"attachments : {attachments}")

    clean = make_clean_leveldb(leveldb)
    try:
        wrap = idb.WrappedIndexDB(clean, blob)

        # locate the account data + attachment databases by name prefix
        offline_db = attach_db = None
        for meta in wrap.database_ids:
            nm = meta.name or ""
            if nm.startswith(OFFLINE_DB_PREFIX):
                offline_db = wrap[meta.dbid_no]
            elif nm.startswith(ATTACH_DB_PREFIX):
                attach_db = wrap[meta.dbid_no]
        if offline_db is None:
            log(f"ERROR: no '{OFFLINE_DB_PREFIX}*' database found in store.")
            sys.exit(4)
        log(f"account db  : {offline_db.name}")

        out_eml = os.path.join(args.out, "EML")
        os.makedirs(out_eml, exist_ok=True)

        # ---- folder tree ----
        folders = {}
        g, _ = iter_store(offline_db, "folders")
        for rec in g:
            v = rec.value
            folders[v["id"]] = {
                "name": dec(v.get("displayName")),
                "parent": v.get("parentFolderId"),
                "dft": v.get("distinguishedFolderType"),
                "serverCount": v.get("totalMessageCount"),
            }
        roots = {"Top of Information Store", "msgfolderroot", "IPM_SUBTREE"}

        def fpath(fid, seen=None):
            seen = seen or set()
            f = folders.get(fid)
            if not f or fid in seen:
                return []
            seen.add(fid)
            nm = f["name"] or f["dft"] or "Folder"
            if nm in roots:
                return []
            parent = folders.get(f["parent"]) if f["parent"] else None
            if not parent or parent["name"] in roots or parent["dft"] == "msgfolderroot":
                return [nm]
            return fpath(f["parent"], seen) + [nm]

        folder_paths = {fid: "/".join(fpath(fid)) for fid in folders}

        # ---- attachment store (attachmentId -> fileName,fileSize) ----
        store = {}
        if attach_db is not None:
            try:
                g, _ = iter_store(attach_db, "attachments")
                for rec in g:
                    v = rec.value
                    aid = v.get("attachmentId")
                    if aid:
                        store[aid] = (dec(v.get("fileName")), v.get("fileSize"))
            except Exception as e:
                log(f"  attachment store unreadable: {e}")

        # ---- disk attachment index ----
        by_namesize, by_name, by_size = {}, collections.defaultdict(list), collections.defaultdict(list)
        if attachments:
            for root, _, files in os.walk(attachments):
                for fn in files:
                    p = os.path.join(root, fn)
                    sz = os.path.getsize(p)
                    if sz == 0:
                        continue
                    by_namesize.setdefault((fn, sz), p)
                    by_name[fn].append(p)
                    by_size[sz].append(p)

        def find_file(name, size, aid=None, inline=False):
            sname, ssize = store.get(aid, (None, None))
            names = [n for n in (name, sname) if n]
            sizes = [s for s in (ssize, size) if s]
            for n in names:
                for s in sizes:
                    if (n, s) in by_namesize:
                        return by_namesize[(n, s)]
            for n in names:
                lst = by_name.get(n)
                if lst and len(lst) == 1:
                    return lst[0]
            if not inline:
                for s in sizes:
                    cand = by_size.get(s, [])
                    if cand:
                        return cand[0]
            for n in names:
                lst = by_name.get(n)
                if lst:
                    return lst[0]
            return None

        # ---- messages ----
        manifest_msgs = []
        stats = {"messages": 0, "with_body": 0, "att_embedded": 0, "att_missing_doc": 0, "att_missing_inline": 0}
        per_folder = collections.Counter()
        g, mstats = iter_store(offline_db, "messageBodies")
        seen = set()
        for rec in g:
            v = rec.value
            if v.get("__typename") != "Message":
                continue
            mid = v.get("id") or (v.get("ItemId") or {}).get("Id")
            if mid in seen:
                continue
            seen.add(mid)
            stats["messages"] += 1

            m = EmailMessage()
            frm = addr(v.get("From")) or addr(v.get("Sender"))
            if frm:
                m["From"] = frm
            to, cc = addrs(v.get("ToRecipients")), addrs(v.get("CcRecipients"))
            if to:
                m["To"] = to
            if cc:
                m["Cc"] = cc
            subj = dec(v.get("Subject"))
            m["Subject"] = subj
            dt = v.get("DateTimeSent") or v.get("DateTimeReceived") or v.get("DateTimeCreated")
            date_iso = None
            if dt:
                try:
                    d = datetime.fromisoformat(dt)
                    date_iso = d.isoformat()
                    m["Date"] = email.utils.format_datetime(d)
                except Exception:
                    pass
            imid = v.get("InternetMessageId")
            if imid:
                m["Message-ID"] = imid
            for h, hh in (("References", "References"), ("InReplyTo", "In-Reply-To")):
                hv = v.get(h)
                if hv:
                    m[hh] = hv if isinstance(hv, str) else dec(hv)

            body, btype = None, "html"
            for bf in ("NormalizedBody", "UniqueBody", "Body"):
                bc = v.get(bf)
                if isinstance(bc, dict) and bc.get("Value") is not None:
                    body = dec(bc.get("Value"))
                    btype = (bc.get("BodyType") or "HTML").lower()
                    break
            if body is None:
                body, btype = dec(v.get("Preview")), "text"
            if btype == "html":
                m.set_content("Diese Nachricht enthält HTML-Inhalt (siehe HTML-Teil).")
                m.add_alternative(body or "", subtype="html")
            else:
                m.set_content(body or "")
            if body:
                stats["with_body"] += 1

            for a in (v.get("Attachments") or []):
                if not isinstance(a, dict):
                    continue
                name = dec(a.get("Name"))
                ctype = a.get("ContentType") or "application/octet-stream"
                size = a.get("Size")
                cid = a.get("ContentId")
                inline = bool(a.get("IsInline"))
                aid = (a.get("AttachmentId") or {}).get("Id")
                p = find_file(name, size, aid=aid, inline=inline)
                if p:
                    with open(p, "rb") as fh:
                        data = fh.read()
                    maintype, _, subtype = ctype.partition("/")
                    if not subtype:
                        maintype, subtype = "application", "octet-stream"
                    kw = {}
                    if inline:
                        kw["disposition"] = "inline"
                        if cid:
                            kw["cid"] = cid if str(cid).startswith("<") else f"<{cid}>"
                    try:
                        m.add_attachment(data, maintype=maintype, subtype=subtype,
                                         filename=name or "attachment", **kw)
                        stats["att_embedded"] += 1
                    except Exception:
                        stats["att_missing_doc"] += 1
                else:
                    stats["att_missing_inline" if inline else "att_missing_doc"] += 1

            pfid = (v.get("ParentFolderId") or {}).get("Id")
            fpath_str = folder_paths.get(pfid) or "_Unknown"
            per_folder[fpath_str] += 1
            d = os.path.join(out_eml, *[safe(x) for x in fpath_str.split("/")]) if fpath_str != "_Unknown" else os.path.join(out_eml, "_Unknown")
            os.makedirs(d, exist_ok=True)
            dtp = ""
            if date_iso:
                try:
                    dtp = datetime.fromisoformat(date_iso).strftime("%Y%m%d_%H%M%S_")
                except Exception:
                    pass
            idsuf = re.sub(r"[^A-Za-z0-9]", "", (mid or "x"))[-8:] or "x"
            rel = os.path.join(*[safe(x) for x in fpath_str.split("/")], safe(f"{dtp}{subj or 'kein_Betreff'}")[:90] + "_" + idsuf + ".eml") if fpath_str != "_Unknown" else os.path.join("_Unknown", idsuf + ".eml")
            with open(os.path.join(out_eml, rel), "wb") as fh:
                fh.write(bytes(m))

            manifest_msgs.append({
                "file": rel.replace(os.sep, "/"),
                "folder": fpath_str,
                "internetMessageId": imid,
                "date": date_iso,
                "isRead": bool(v.get("IsRead")),
                "isDraft": bool(v.get("IsDraft")),
                "hasAttachments": bool(v.get("Attachments")),
            })
            if args.json_progress and stats["messages"] % 25 == 0:
                print(json.dumps({"progress": stats["messages"]}), flush=True)

        manifest = {
            "account": offline_db.name.replace(OFFLINE_DB_PREFIX, ""),
            "generated": datetime.now().isoformat(timespec="seconds"),
            "folders": [
                {"id": fid, "name": f["name"], "path": folder_paths.get(fid),
                 "serverCount": f["serverCount"]}
                for fid, f in folders.items()
            ],
            "messages": manifest_msgs,
            "stats": {**stats, "per_folder": dict(per_folder),
                      "messageBodies_records": mstats},
        }
        with open(os.path.join(args.out, "manifest.json"), "w", encoding="utf-8") as fh:
            json.dump(manifest, fh, ensure_ascii=False, indent=1)

        log("---- done ----")
        log(json.dumps({**stats, "per_folder": dict(per_folder)}, ensure_ascii=False, indent=1))
        print(json.dumps({"done": True, "messages": stats["messages"],
                          "out": args.out}), flush=True)
    finally:
        shutil.rmtree(clean, ignore_errors=True)


if __name__ == "__main__":
    main()
