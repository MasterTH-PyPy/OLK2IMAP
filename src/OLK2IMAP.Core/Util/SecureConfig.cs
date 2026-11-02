using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OLK2IMAP.Core.Util;

/// <summary>
/// AES-256-GCM encrypted settings blob (IMAP/EWS credentials etc.), key derived
/// from a user passphrase via PBKDF2. Same pattern as the other SonOG admin tools.
/// File layout: [16B salt][12B nonce][16B tag][ciphertext].
/// </summary>
public static class SecureConfig
{
    private const int SaltSize = 16, NonceSize = 12, TagSize = 16, KeySize = 32;
    private const int Iterations = 200_000;

    public static void Save<T>(string path, T value, string passphrase)
    {
        var json = JsonSerializer.SerializeToUtf8Bytes(value);
        var salt = RandomNumberGenerator.GetBytes(SaltSize);
        var nonce = RandomNumberGenerator.GetBytes(NonceSize);
        var key = DeriveKey(passphrase, salt);
        var cipher = new byte[json.Length];
        var tag = new byte[TagSize];
        using (var aes = new AesGcm(key, TagSize))
            aes.Encrypt(nonce, json, cipher, tag);

        using var fs = File.Create(path);
        fs.Write(salt); fs.Write(nonce); fs.Write(tag); fs.Write(cipher);
    }

    public static T Load<T>(string path, string passphrase)
    {
        var all = File.ReadAllBytes(path);
        var salt = all.AsSpan(0, SaltSize).ToArray();
        var nonce = all.AsSpan(SaltSize, NonceSize).ToArray();
        var tag = all.AsSpan(SaltSize + NonceSize, TagSize).ToArray();
        var cipher = all.AsSpan(SaltSize + NonceSize + TagSize).ToArray();
        var key = DeriveKey(passphrase, salt);
        var plain = new byte[cipher.Length];
        using (var aes = new AesGcm(key, TagSize))
            aes.Decrypt(nonce, cipher, tag, plain);
        return JsonSerializer.Deserialize<T>(plain)
               ?? throw new InvalidDataException("Config leer/ungültig.");
    }

    private static byte[] DeriveKey(string passphrase, byte[] salt)
        => Rfc2898DeriveBytes.Pbkdf2(Encoding.UTF8.GetBytes(passphrase), salt, Iterations,
                                     HashAlgorithmName.SHA256, KeySize);
}
