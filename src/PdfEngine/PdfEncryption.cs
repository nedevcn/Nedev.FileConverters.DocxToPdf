using System.Security.Cryptography;
using System.Text;

namespace Nedev.DocxToPdf.PdfEngine;

public class PdfEncryption
{
    private readonly byte[] _userPassword;
    private readonly byte[] _ownerPassword;
    private readonly int _permissions;
    private readonly byte[] _encryptionKey;
    
    public static readonly int PRINT = 4;
    public static readonly int MODIFY = 8;
    public static readonly int COPY = 16;
    public static readonly int FILL_FORM = 32;

    public PdfEncryption(string? userPassword, string? ownerPassword, int permissions)
    {
        _userPassword = Encoding.UTF8.GetBytes(userPassword ?? "");
        _ownerPassword = string.IsNullOrEmpty(ownerPassword) 
            ? GenerateRandomPassword() 
            : Encoding.UTF8.GetBytes(ownerPassword);
        _permissions = permissions;
        
        _encryptionKey = ComputeEncryptionKey(_userPassword, _ownerPassword, _permissions);
    }

    private static byte[] GenerateRandomPassword()
    {
        var random = new Random();
        const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        var password = new string(Enumerable.Range(0, 32).Select(_ => chars[random.Next(chars.Length)]).ToArray());
        return Encoding.UTF8.GetBytes(password);
    }

    private static byte[] ComputeEncryptionKey(byte[] userPwd, byte[] ownerPwd, int permissions)
    {
        var hash = new List<byte>();
        hash.AddRange(userPwd);
        hash.AddRange(ownerPwd);
        hash.Add((byte)(permissions & 0xFF));
        hash.Add((byte)((permissions >> 8) & 0xFF));
        hash.AddRange(Encoding.ASCII.GetBytes("Nedev.DocxToPdf"));
        
        using var sha256 = SHA256.Create();
        return sha256.ComputeHash(hash.ToArray()).Take(16).ToArray();
    }

    public byte[] GetUserPasswordBytes()
    {
        return _userPassword;
    }

    public byte[] GetOwnerPasswordBytes()
    {
        return _ownerPassword;
    }

    public int GetPermissions()
    {
        return _permissions;
    }

    public byte[] GetEncryptionKey()
    {
        return _encryptionKey;
    }

    public string ComputeOHash()
    {
        using var sha256 = SHA256.Create();
        var hash = sha256.ComputeHash(_ownerPassword);
        return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant().Substring(0, 32);
    }

    public string ComputeUHash()
    {
        var rc4Key = _encryptionKey.Take(16).ToArray();
        var rc4 = new RC4Cipher(rc4Key);
        var paddedPwd = _userPassword.ToList();
        while (paddedPwd.Count < 32) paddedPwd.Add(0);
        var encrypted = rc4.Encrypt(paddedPwd.ToArray());
        return BitConverter.ToString(encrypted).Replace("-", "").ToLowerInvariant();
    }
}

public class RC4Cipher
{
    private readonly byte[] _s;
    private int _i;
    private int _j;

    public RC4Cipher(byte[] key)
    {
        _s = new byte[256];
        for (int i = 0; i < 256; i++)
            _s[i] = (byte)i;

        int j = 0;
        for (int i = 0; i < 256; i++)
        {
            j = (j + _s[i] + key[i % key.Length]) % 256;
            (_s[i], _s[j]) = (_s[j], _s[i]);
        }
        _i = 0;
        _j = 0;
    }

    public byte[] Encrypt(byte[] data)
    {
        var result = new byte[data.Length];
        for (int n = 0; n < data.Length; n++)
        {
            _i = (_i + 1) % 256;
            _j = (_j + _s[_i]) % 256;
            (_s[_i], _s[_j]) = (_s[_j], _s[_i]);
            int k = _s[(_s[_i] + _s[_j]) % 256];
            result[n] = (byte)(data[n] ^ k);
        }
        return result;
    }
}
