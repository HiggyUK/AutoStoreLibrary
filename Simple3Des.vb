' ***********************************************************************
' Assembly         : AutoStoreLibrary
' Author           : John Campbell-Higgens
' Created          : 04-Jan-2021
'
' Last Modified By : John Campbell-Higgens
' Last Modified On : 06-Jan-2021
' ***********************************************************************
' <copyright file="Simple3Des.vb" company="John Campbell-Higgens (Kofax UK Ltd)">
'     Copyright © 2020 John Campbell-Higgens (Kofax UK Ltd)
' </copyright>
' <summary></summary>
' ***********************************************************************

Imports System.Security.Cryptography
''' <summary>
''' Simple Triple Des encryption function libary. Used by the <see cref="Tools" />function class of the AutoStoreLibrary
''' </summary>
Public NotInheritable Class Simple3Des

    ''' <summary>
    ''' The triple DES
    ''' </summary>
    Private TripleDes As New TripleDESCryptoServiceProvider

    ''' <summary>
    ''' Internal function of the Simple3Des library in AutoStoreLibrary to truncate and has the provided string
    ''' </summary>
    ''' <param name="key">The key.</param>
    ''' <param name="length">The length.</param>
    ''' <returns>System.Byte().</returns>
    Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()

        Dim sha1 As New SHA1CryptoServiceProvider

        ' Hash the key.
        Dim keyBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)

        ' Truncate or pad the hash.
        ReDim Preserve hash(length - 1)
        Return hash
    End Function

    ''' <summary>
    ''' Creates a new Key
    ''' </summary>
    ''' <param name="key">The key.</param>
    Public Sub New(ByVal key As String)
        ' Initialize the crypto provider.
        TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub


    ''' <summary>
    ''' Encodes a string of text with TripleDes encryption
    ''' </summary>
    ''' <param name="plaintext">Text to be encoded</param>
    ''' <returns>System.String.</returns>
    Public Function EncryptData(ByVal plaintext As String) As String

        Dim encodedString As String = ""

        ' Convert the plaintext string to a byte array.
        Dim plaintextBytes() As Byte =
            System.Text.Encoding.Unicode.GetBytes(plaintext)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the encoder to write to the stream.
        Dim encStream As New CryptoStream(ms,
            TripleDes.CreateEncryptor(),
            System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
        encStream.FlushFinalBlock()

        ' Convert the encrypted stream to a printable string.
        encodedString = Convert.ToBase64String(ms.ToArray)
        Return encodedString

    End Function

    ''' <summary>
    ''' Decodes a string of TripleDes encrypted text.
    ''' </summary>
    ''' <param name="encryptedtext">Text to be decoded</param>
    ''' <returns>System.String.</returns>
    Public Function DecryptData(ByVal encryptedtext As String) As String

        ' Convert the encrypted text string to a byte array.
        Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the decoder to write to the stream.
        Dim decStream As New CryptoStream(ms,
            TripleDes.CreateDecryptor(),
            System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
        decStream.FlushFinalBlock()

        ' Convert the plaintext stream to a string.
        Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
    End Function
End Class
