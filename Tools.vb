' ***********************************************************************
' Assembly         : AutoStoreLibrary
' Author           : John Campbell-Higgens
' Created          : 31-Dec-2020
'
' Last Modified By : John Campbell-Higgens
' Last Modified On : 06-Jan-2021
' ***********************************************************************
' <copyright file="Tools.vb" company="John Campbell-Higgens (Kofax UK Ltd)">
'     Copyright © 2020 John Campbell-Higgens (Kofax UK Ltd)
' </copyright>
' <summary></summary>
' ***********************************************************************

Imports System.IO
Imports System.Security.Cryptography

''' <summary>
''' Function Library of useful Tools that can be used within AutoStore Scripts.
''' </summary>
Public Class Tools

    ''' <summary>
    ''' Returns the version of the Tools Library elements of this AutoStore Library
    ''' </summary>
    ''' <returns>Version String</returns>
    Public Shared Function GetVersion() As String

        Return "Version 1.2.0.4"

    End Function
    ''' <summary>
    ''' <para>Encrypts a plain text string using TripleDes
    ''' encryption.</para>
    ''' <para>The encoded data can only be Decoded using the <see cref="DecodeData" />method in this library.</para>
    ''' </summary>
    ''' <param name="plaintext">Text to encrypt</param>
    ''' <returns>Encrypted text</returns>
    Public Shared Function EncodeData(ByVal plaintext As String) As String

        Dim wrapper As New Simple3Des("Kofax2021!")
        Dim cipherText As String = wrapper.EncryptData(plaintext)

        Return cipherText

    End Function
    ''' <summary>
    ''' <para>
    ''' Decrypts data which has been encrypted using TripleDes</para>
    ''' <para>
    '''   <font color="#333333">Will only decode text which has been encrypted using the <see cref="EncodeData" /> method</font>
    ''' </para>
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
    ''' <para>
    ''' Saves an a string to a file based on the username passed. </para>
    ''' <para>File is named after the username with the extension .PSW
    ''' and is stored in the location provided.</para>
    ''' <para>The calling code must have access to the cache location. </para>
    ''' </summary>
    ''' <param name="encryptedPassword">Encrypted password to stored</param>
    ''' <param name="userName">Username of the password</param>
    ''' <param name="cacheLocation">Location of the password cache.  This is a folder on the KCS AutoStore server</param>
    ''' <returns><br /></returns>
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
    ''' <para>
    ''' Get the stored text from the saved password file. The password file is a text file with a single line of text.</para>
    ''' <para>The password file name is based on the username with the extension of .PSW.
    ''' The calling code must have access to the cache location. </para>
    ''' <para>If the file does not exist then a blank password is returned. </para>
    ''' </summary>
    ''' <param name="userName">Username to retrieve string for</param>
    ''' <param name="cacheLocation">Location of the password cache.  This should be a directory on the KCS AutoStore Server</param>
    ''' <returns>String</returns>
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


