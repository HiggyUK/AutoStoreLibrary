Imports System.IO
Imports System.Net
Imports System.Collections.Generic
Imports System.Linq
Imports System.Net.Cache
Imports System.Net.Http
Imports DocuWare.Platform.ServerClient
Imports Newtonsoft.Json


Public Class DocuWare

    Shared uri As Uri = New Uri("http://chw-dw-03/docuware/platform")


    ''' <summary>
    ''' Returns the version of the DocuWare Library elements of this AutoStore Library
    ''' </summary>
    ''' <returns>Version String</returns>
    Public Shared Function GetVersion() As String

        Return "Version 1.1.0.3"

    End Function

    ''' <summary>
    ''' Function to return a list of Organisations that the user has access to.
    ''' </summary>
    ''' <param name="uri">The path to the Document Instance, either on premise or cloud.  This is usually in the format of http://[server/cloud]/docuware/platform</param>
    ''' <param name="userName">Username within DocuWare</param>
    ''' <param name="userPassword">Password of user in Docuware</param>
    ''' <returns>List of Organisation Names that the user has access too.  Which may only be a single return</returns>
    Public Shared Function GetOrganisations(uri As Uri, userName As String, userPassword As String) As String

        Dim docuWareConnection As ServiceConnection
        Dim docuWareOrg As Organization
        Dim orgCount As Integer
        Dim organisationName As String = ""

        Try
            docuWareConnection = ServiceConnection.Create(uri, userName, userPassword)
            If docuWareConnection.Organizations.Count > 0 Then
                For orgCount = 0 To docuWareConnection.Organizations.Count - 1
                    docuWareOrg = docuWareConnection.Organizations(orgCount)
                    organisationName = "|" + docuWareOrg.Name
                Next
            End If
            docuWareConnection.Disconnect()

        Catch ex As Exception
            organisationName = ex.Message
        End Try

        If Strings.Left(organisationName, 1) = "|" Then
            organisationName = Strings.Right(organisationName, Len(organisationName) - 1)
        End If

        Return organisationName

    End Function
    ''' <summary>
    ''' Function to return the name of the primary organisation that the user has access too, assuming they have access to it.
    ''' </summary>
    ''' <param name="uri">The path to the Document Instance, either on premise or cloud.  This is usually in the format of http://[server/cloud]/docuware/platform</param>
    ''' <param name="userName">DocuWare username</param>
    ''' <param name="userPassword">Password of the DocuWare user</param>
    ''' <param name="organisation">Organistion to check and return</param>
    ''' <returns>Organisation Name</returns>
    Public Shared Function GetOrganisation(uri As Uri, userName As String, userPassword As String, organisation As String) As String

        Dim docuWareConnection As ServiceConnection
        Dim docuWareOrg As Organization
        Dim organisationName As String = ""

        docuWareConnection = ServiceConnection.Create(uri, userName, userPassword, organization:=organisation)
        docuWareOrg = docuWareConnection.Organizations(0)

        organisationName = docuWareOrg.Name
        docuWareConnection.Disconnect()

        Return organisationName

    End Function
    ''' <summary>
    ''' Function to return a list of File Cabinets that the user has access to in the organisation specified.
    ''' </summary>
    ''' <param name="uri">The path to the Document Instance, either on premise or cloud.  This is usually in the format of http://[server/cloud]/docuware/platform</param>
    ''' <param name="userName">DocuWare UserName</param>
    ''' <param name="userPassword">DocuWare user password</param>
    ''' <param name="organisation">Organisation name to return File Cabinets for</param>
    ''' <returns>List of File Cabinets delimied with the | symbol</returns>
    Public Shared Function GetFileCabinet(uri As Uri, userName As String, userPassword As String, organisation As String) As String

        Dim docuWareConnection As ServiceConnection
        Dim docuWareOrg As Organization
        Dim docuWareFileCabinets As List(Of FileCabinet)
        Dim fileCabinetList As String = ""

        docuWareConnection = ServiceConnection.Create(uri, userName, userPassword, organization:=organisation)
        docuWareOrg = docuWareConnection.Organizations(0)

        docuWareFileCabinets = docuWareOrg.GetFileCabinetsFromFilecabinetsRelation().FileCabinet

        For Each fc In docuWareFileCabinets
            If fc.IsBasket = False Then
                fileCabinetList = fileCabinetList + "|" + fc.Name
            End If
        Next

        If Strings.Left(fileCabinetList, 1) = "|" Then
            fileCabinetList = Strings.Right(fileCabinetList, Len(fileCabinetList) - 1)
        End If

        docuWareConnection.Disconnect()

        Return fileCabinetList

    End Function
    ''' <summary>
    ''' Function to return the list of Baskets that the user has access to in the specified organisation
    ''' </summary>
    ''' <param name="uri">The path to the Document Instance, either on premise or cloud.  This is usually in the format of http://[server/cloud]/docuware/platform</param>
    ''' <param name="userName">Docuware UserName</param>
    ''' <param name="userPassword">Docuware Userpassword</param>
    ''' <param name="organisation">Organisation Name</param>
    ''' <returns>List of Baskets that the user has access to delimited with | character</returns>
    Public Shared Function GetBaskets(uri As Uri, userName As String, userPassword As String, organisation As String) As String

        Dim docuWareConnection As ServiceConnection
        Dim docuWareOrg As Organization
        Dim docuWareFileCabinets As List(Of FileCabinet)
        Dim fileCabinetList As String = ""

        docuWareConnection = ServiceConnection.Create(uri, userName, userPassword, organization:=organisation)
        docuWareOrg = docuWareConnection.Organizations(0)

        docuWareFileCabinets = docuWareOrg.GetFileCabinetsFromFilecabinetsRelation().FileCabinet

        For Each fc In docuWareFileCabinets
            If fc.IsBasket = True Then
                fileCabinetList = fileCabinetList + "|" + fc.Name
            End If
        Next

        If Strings.Left(fileCabinetList, 1) = "|" Then
            fileCabinetList = Strings.Right(fileCabinetList, Len(fileCabinetList) - 1)
        End If

        docuWareConnection.Disconnect()

        Return fileCabinetList

    End Function

    ''' <summary>
    ''' Function to return a complete list of fields
    ''' </summary>
    ''' <param name="uri">The path to the Document Instance, either on premise or cloud.  This is usually in the format of http://[server/cloud]/docuware/platform</param>
    ''' <param name="userName">Docuware UserName</param>
    ''' <param name="userPassword">Docuware Userpassword</param>
    ''' <param name="organisation">Organisation Name</param>
    ''' <param name="cabinet">Cabinet Name</param>
    ''' <returns>A list of fields which are delimited with | characters.  Each field will be returned with the following fields
    ''' 
    ''' Field DB Name
    ''' Field Display Name
    ''' Field Type
    ''' Field Scope
    ''' 
    ''' Each field will be seperated with ||
    ''' 
    ''' For example if the cabinet has 3 fields they will be returned as 
    ''' fieldname1|Field 1|0|1||fieldname2|Field 2|0|1||fieldname3|Field 3|0|1||
    ''' 
    ''' Scope is currently 0 or 1 
    ''' O - System Field
    ''' 1 - User Field
    ''' 
    ''' </returns>
    Public Shared Function GetFields(uri As Uri, userName As String, userPassword As String, organisation As String, cabinet As String) As String

        Dim fieldList As String = ""
        Dim docuWareConnection As ServiceConnection
        Dim docuWareOrg As Organization
        Dim fileCabinet As FileCabinet
        Dim fileCabinets As List(Of FileCabinet)
        Dim fileCabinetField As FileCabinetField
        Dim fileCabinetFields As List(Of FileCabinetField)
        Dim fieldNameDB As String
        Dim fieldNameDisplay As String
        Dim fieldNameType As String
        Dim fieldNameScope As String
        docuWareConnection = ServiceConnection.Create(uri, userName, userPassword, organization:=organisation)
        docuWareOrg = docuWareConnection.Organizations(0)



        fileCabinets = docuWareOrg.GetFileCabinetsFromFilecabinetsRelation().FileCabinet

        For Each fileCabinet In fileCabinets
            If String.Compare(fileCabinet.Name, cabinet, True) = 0 Then
                fileCabinet = fileCabinet.GetFileCabinetFromSelfRelation()
                fileCabinetFields = fileCabinet.Fields
                For Each fileCabinetField In fileCabinetFields
                    fieldNameDB = fileCabinetField.DBFieldName
                    fieldNameDisplay = fileCabinetField.DisplayName
                    fieldNameType = fileCabinetField.DWFieldType
                    fieldNameScope = fileCabinetField.Scope
                    fieldList = fieldList + fieldNameDB + "|" + fieldNameDisplay + "|" + fieldNameType + "|" + fieldNameScope + "||"
                Next
            End If
        Next

        docuWareConnection.Disconnect()

        Return fieldList

    End Function
    Public Shared Function Connect() As ServiceConnection
        Return ServiceConnection.Create(uri, "admin", "admin")
    End Function

    Public Shared Function ConnectWithUserAgent() As ServiceConnection
        Return ServiceConnection.Create(uri, "admin", "admin", userAgent:=New System.Net.Http.Headers.ProductInfoHeaderValue() {New System.Net.Http.Headers.ProductInfoHeaderValue("DocuWare+.NET+API+Test+Client", "1.0")})
    End Function

    Public Shared Function ConnectWithOrg() As ServiceConnection
        Return ServiceConnection.Create(uri, "admin", "admin", organization:="Peters Engineering")
    End Function
    Public Shared Function ConnectWithCaching() As ServiceConnection
        Dim handler = New WebRequestHandler() With {
            .CachePolicy = New RequestCachePolicy(RequestCacheLevel.[Default])
        }
        Return ServiceConnection.Create(uri, "admin", "admin", httpClientHandler:=handler)
    End Function

    Public Shared Function ConnectWithNTLM() As ServiceConnection
        Return ServiceConnection.CreateWithWindowsAuthentication(uri, "Administrator", "admin")
    End Function

    Public Shared Function ConnectWithDefaultUser() As ServiceConnection
        Return ServiceConnection.CreateWithWindowsAuthentication(uri, System.Net.CredentialCache.DefaultCredentials)
    End Function

    Public Shared Function ConnectWithDisplayLanguageAndFormatCulture() As ServiceConnection
        Dim handler = New WebRequestHandler() With {
            .CookieContainer = New CookieContainer(),
            .UseCookies = True
        }
        handler.CookieContainer.Add(New Cookie("DWFormatCulture", "de", "/", "localhost"))
        handler.CookieContainer.Add(New Cookie("DWLanguage", "de", "/", "localhost"))
        Return ServiceConnection.Create(uri, "admin", "admin", httpClientHandler:=handler)
    End Function

End Class
