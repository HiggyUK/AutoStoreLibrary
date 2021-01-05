Imports System.IO
Imports System.Security.Cryptography

Public Class Tools

    ''' <summary>
    ''' Returns the version of the Tools Library elements of this AutoStore Library
    ''' </summary>
    ''' <returns>Version String</returns>
    Public Shared Function GetVersion() As String

        Return "Version 1.1.0.4"

    End Function
    ''' <summary>
    ''' Encrypts a plain text string using TripleDes
    ''' </summary>
    ''' <param name="plaintext">Text to encrypt</param>
    ''' <returns>Encrypted text</returns>
    Public Shared Function EncodeData(ByVal plaintext As String) As String

        Dim wrapper As New Simple3Des("Kofax2021!")
        Dim cipherText As String = wrapper.EncryptData(plaintext)

        Return cipherText

    End Function
    ''' <summary>
    ''' Decrypts data which has been encrypted using TripleDes
    ''' </summary>
    ''' <param name="encryptedtext">Encrypted text to be decrypted</param>
    ''' <returns>Plain Text</returns>
    Public Shared Function DecodeData(ByVal encryptedtext As String) As String

        Dim plaintext As String
        Dim wrapper As New Simple3Des("Kofax2021!")
        Try
            plaintext = wrapper.DecryptData(encryptedtext)
        Catch ex As Exception
            plaintext = "Error: Unable to Decode"
        End Try

        Return plaintext

    End Function

    ''' <summary>
    ''' Saves an encrypted password to a file based on the username.  File is named after the username with the extension .PSW
    ''' </summary>
    ''' <param name="encryptedPassword">Encrypted password to stored</param>
    ''' <param name="userName">Username of the password</param>
    ''' <param name="cacheLocation">Location of the password cache.  This is a folder on the KCS AutoStore server</param>
    ''' <returns></returns>
    Public Shared Function SavePassword(ByVal encryptedPassword As String, ByVal userName As String, ByVal cacheLocation As String) As String

        Dim passwordSaved As String = "OK"

        ' Try Directory
        Try
            If Not Directory.Exists(cacheLocation) Then
                Directory.CreateDirectory(cacheLocation)
            End If
            Dim cacheFile As String = Path.Combine(cacheLocation, userName + ".psw")
            File.Delete(cacheFile)
            Dim cacheWrite As StreamWriter
            cacheWrite = My.Computer.FileSystem.OpenTextFileWriter(cacheFile, False)
            cacheWrite.WriteLine(encryptedPassword)
            cacheWrite.Close()

        Catch ex As Exception
            passwordSaved = ex.Message
        End Try

        Return passwordSaved

    End Function

    ''' <summary>
    ''' Get the stored encrypted password from the saved password file. The password file name is based on the username with the extension of .PSW.
    ''' </summary>
    ''' <param name="userName">Username to retrieve encryped password for</param>
    ''' <param name="cacheLocation">Location of the password cache.  This should be a directory on the KCS AutoStore Server</param>
    ''' <returns></returns>
    Public Shared Function GetPassword(ByVal userName As String, ByVal cacheLocation As String) As String

        Dim encryptedPassword As String = ""

        Try
            Dim cacheFile As String = Path.Combine(cacheLocation, userName + ".psw")
            If File.Exists(cacheFile) Then

                Dim cacheRead As StreamReader
                cacheRead = My.Computer.FileSystem.OpenTextFileReader(cacheFile)
                encryptedPassword = cacheRead.ReadLine()
                cacheRead.Close()
            Else
                encryptedPassword = ""
            End If
        Catch ex As Exception
            encryptedPassword = "ERROR:" + ex.Message
        End Try

        Return encryptedPassword

    End Function





End Class


