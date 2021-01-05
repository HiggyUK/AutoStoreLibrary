Imports System.IO
Imports System.Net
Imports Newtonsoft.Json

Public Class SignDoc

    ''' <summary>
    ''' Retrieve the version of the AutoStoreLibrary for SignDoc
    ''' </summary>
    ''' <returns>Verison string</returns>
    Public Shared Function GetVersion() As String

        Return "Version 1.1.0.4"


    End Function

    ''' <summary>
    ''' Returns the Authentication Token for the user and account at the SignDoc Server. This function should be used before all other functions
    ''' </summary>
    ''' <param name="serverAddress">URL address of the SignDoc Server</param>
    ''' <param name="userName">SignDoc Username</param>
    ''' <param name="userPassword">SignDoc Users password</param>
    ''' <param name="accountID">SignDoc Account ID</param>
    ''' <returns>Authentiction Token</returns>
    Public Shared Function GetAuthenticationToken(ByVal serverAddress, ByVal userName, ByVal userPassword, ByVal accountID) As String

        Dim AuthenticationToken As String = ""

        Dim headers As WebHeaderCollection
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = (SecurityProtocolType.Tls12)

        userName = userName.Replace("@", "%40")
        Dim serverURL As String = serverAddress + "/rest/v7/users/authentication?credentials=" + userName + "&accountid=" + accountID + "&password=" & userPassword & "&usedefaultaccount=false"

        Try
            request = WebRequest.Create(serverURL)
            request.ContentType = "application/json"
            request.Method = "POST"

            response = request.GetResponse

            If response.StatusCode = System.Net.HttpStatusCode.OK Then
                headers = response.Headers

                For header As Integer = 0 To headers.Count - 1
                    If headers.Keys(header) = "X-AUTH-TOKEN" Then
                        AuthenticationToken = headers(header)
                    End If
                Next

            Else
                AuthenticationToken = response.StatusCode & " - " & response.StatusDescription
            End If
        Catch ex As Exception
            AuthenticationToken = ex.Message & " Call: " & serverURL
        End Try

        Return AuthenticationToken

    End Function

    ''' <summary>
    ''' Function to create a SignDoc package.  This creates the initial package that documents and signature can then be added to.
    ''' </summary>
    ''' <param name="packageName">The name of the SignDoc Package</param>
    ''' <param name="packageType">the SignDoc package type - PACKAGE is currently the only supported option </param>
    ''' <param name="AuthenticationToken">Authentication token from the GetAuthenticationToken function</param>
    ''' <param name="serverAddress">URL address of the SignDoc Sever</param>
    ''' <returns>PackageID of the created SignDoc Package</returns>
    Public Shared Function CreatePackage(ByVal packageName, ByVal packageType, ByVal AuthenticationToken, ByVal serverAddress) As String
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = (SecurityProtocolType.Tls12)


        Dim packageID As String = ""
        Dim packageURL As String
        Dim stream As Stream
        Dim streamRead As StreamReader
        Dim jsonReturn As String

        Dim NewPackage As New SignDocPackage

        NewPackage.name = packageName
        NewPackage.type = packageType

        Dim jsonData As String = JsonConvert.SerializeObject(NewPackage, Formatting.Indented)

        Dim serverURL As String = serverAddress + "/rest/v7/package?schedule=false&clean_fields=false&delete_existing=false&autoprepare=false&resolution=72"
        Try
            request = WebRequest.Create(serverURL)
            request.ContentType = "application/json"
            request.Accept = "application/json"
            request.Method = "POST"
            request.Headers.Add("X-AUTH-TOKEN", AuthenticationToken)

            Using streamWriter As New StreamWriter(request.GetRequestStream())
                streamWriter.Write(jsonData)
            End Using
            response = request.GetResponse

            If response.StatusCode = 201 Then
                stream = response.GetResponseStream
                streamRead = New StreamReader(stream)
                jsonReturn = streamRead.ReadToEnd
                streamRead.Close()
                stream.Close()

                Dim PackageDetails As SignDocPackageDetails = JsonConvert.DeserializeObject(Of SignDocPackageDetails)(jsonReturn)

                packageID = PackageDetails.id
                packageURL = PackageDetails.url
            End If

        Catch ex As Exception
            packageID = ex.Message
        End Try

        packageID = packageID

        Return packageID

    End Function

    ''' <summary>
    ''' Adds a Signer to the SignDoc Package created by the CreatePackage Function.
    ''' </summary>
    ''' <param name="serverAddress">URL address of the SignDoc Sever</param>
    ''' <param name="authToken">Authentication token from the GetAuthenticationToken function</param>
    ''' <param name="PackageID">Package ID from the CreatePackage function, used to identify the package to add signers to.</param>
    ''' <param name="SignerName">Name of the Signer</param>
    ''' <param name="SignerEmail">Email Address of the Signer</param>
    ''' <param name="SignerType">Type of Signer (Option SIGNER or REVIEWER) </param>
    ''' <returns>Returns the SignerID of the Signer Created</returns>
    Public Shared Function AddSigner(ByVal serverAddress, ByVal authToken, ByVal PackageID, ByVal SignerName, ByVal SignerEmail, ByVal SignerType) As String

        Dim SignerID As String = ""
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse
        Dim stream As Stream
        Dim streamRead As StreamReader
        Dim jsonReturn As String

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = (SecurityProtocolType.Tls12)


        Dim NewSigner As New SignerDetails

        NewSigner.name = SignerName
        NewSigner.email = SignerEmail
        NewSigner.role = SignerType


        Dim jsonData As String = JsonConvert.SerializeObject(NewSigner, Formatting.Indented)

        Dim serverURL As String = serverAddress + "/rest/v7/packages/" & PackageID & "/signer"
        Try
            request = WebRequest.Create(serverURL)
            request.ContentType = "application/json"
            request.Accept = "application/json"
            request.Method = "POST"
            request.Headers.Add("X-AUTH-TOKEN", authToken)

            Using streamWriter As New StreamWriter(request.GetRequestStream())
                streamWriter.Write(jsonData)
            End Using
            response = request.GetResponse

            If response.StatusCode = 201 Then
                stream = response.GetResponseStream
                streamRead = New StreamReader(stream)
                jsonReturn = streamRead.ReadToEnd
                streamRead.Close()
                stream.Close()

                Dim SignerDetails As SignDocSignerDetails = JsonConvert.DeserializeObject(Of SignDocSignerDetails)(jsonReturn)

                SignerID = SignerDetails.id

            End If

        Catch ex As Exception
            SignerID = ex.Message

        End Try
        Return SignerID



    End Function

    ''' <summary>
    ''' Adds a document to the SignDoc signing Package.
    ''' </summary>
    ''' <param name="serverAddress">URL address of the SignDoc Sever</param>
    ''' <param name="authToken">Authentication token from the GetAuthenticationToken function</param>
    ''' <param name="PackageID">Package ID from the CreatePackage function, used to identify the package to add signers to.</param>
    ''' <param name="DocumentFilePath">File Path to the Document to be added to the Package</param>
    ''' <param name="DocumentFileName">File Name of the Document to be added to the Package</param>
    ''' <param name="DocumentFormat">Format of the Document (PDF, TIF)</param>
    ''' <param name="DocumentName">Name of the Document</param>
    ''' <param name="DocumentDescription">Description of the Document</param>
    ''' <param name="DocumentMessage">Message to attach to the document</param>
    ''' <returns>Returns the DocumentID of the added Document </returns>
    Public Shared Function AddDocument(ByVal serverAddress, ByVal authToken, ByVal PackageID, ByVal DocumentFilePath, ByVal DocumentFileName, ByVal DocumentFormat, ByVal DocumentName, ByVal DocumentDescription, ByVal DocumentMessage) As String

        Dim request As HttpWebRequest
        Dim response As HttpWebResponse
        Dim stream As Stream
        Dim streamRead As StreamReader
        Dim jsonReturn As String
        Dim DocumentID As String = ""

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = (SecurityProtocolType.Tls12)


        Dim NewDocument As New DocumentEntry

        NewDocument.content = Convert.ToBase64String(System.IO.File.ReadAllBytes(DocumentFilePath))
        NewDocument.name = DocumentName
        NewDocument.description = DocumentDescription
        NewDocument.documentMessage = DocumentMessage
        NewDocument.fileName = DocumentFileName
        NewDocument.format = DocumentFormat



        Dim jsonData As String = JsonConvert.SerializeObject(NewDocument, Formatting.Indented)

        Dim serverURL As String = serverAddress + "/rest/v7/packages/" & PackageID & "/document?resolution=72&autoprep=false"
        Try
            request = WebRequest.Create(serverURL)
            request.ContentType = "application/json"
            request.Accept = "application/json"
            request.Method = "POST"
            request.Headers.Add("X-AUTH-TOKEN", authToken)

            Using streamWriter As New StreamWriter(request.GetRequestStream())
                streamWriter.Write(jsonData)
            End Using
            response = request.GetResponse

            If response.StatusCode = 201 Then
                stream = response.GetResponseStream
                streamRead = New StreamReader(stream)
                jsonReturn = streamRead.ReadToEnd
                streamRead.Close()
                stream.Close()

                Dim DocumentDetails As DocumentReturn = JsonConvert.DeserializeObject(Of DocumentReturn)(jsonReturn)

                DocumentID = DocumentDetails.id

            End If

        Catch ex As Exception

            DocumentID = ex.Message


        End Try
        Return DocumentID
    End Function


    ''' <summary>
    ''' Adds a Signature Field to a Document for a specified Signer
    ''' </summary>
    ''' <param name="serverAddress">URL address of the SignDoc Sever</param>
    ''' <param name="authToken">Authentication token from the GetAuthenticationToken function</param>
    ''' <param name="PackageID">Package ID from the CreatePackage function</param>
    ''' <param name="DocumentID">Document ID from the AddDocument function</param>
    ''' <param name="fieldName">Name of the Field to Add</param>
    ''' <param name="alternativeName">Alternative Name of the Field to Add</param>
    ''' <param name="SignerID">SignerID of the Signer associated with the field</param>
    ''' <param name="RequiredField">Required Field - True / False</param>
    ''' <param name="ReadOnlyField">ReadOnly Field - True / False</param>
    ''' <param name="SignatureDescription">Description of Field</param>
    ''' <param name="WidgetIndex">Index Number of Field (Integer)</param>
    ''' <param name="WidgetPageNumber">Page Number on the Document for Field (Integer)</param>
    ''' <param name="WidgetTop">Location of the Top of the Field - from Bottom Left of page (Integer)</param>
    ''' <param name="WidgetBottom">Location of the Bottom of the Field - from Bottom Left of page (Integer)</param>
    ''' <param name="WidgetLeft">Location of the Left of the Field - from Bottom Left of Page (Integer)</param>
    ''' <param name="WidgetRight">Location of the Right of the Field - from Bottom Left of Page (Integer)</param>
    ''' <returns>Returns Field ID</returns>
    Public Shared Function AddSignatureField(ByVal serverAddress, ByVal authToken, ByVal PackageID, ByVal DocumentID, ByVal fieldName, ByVal alternativeName, ByVal SignerID, ByVal RequiredField, ByVal ReadOnlyField, ByVal SignatureDescription, ByVal WidgetIndex, ByVal WidgetPageNumber, ByVal WidgetTop, ByVal WidgetBottom, ByVal WidgetLeft, ByVal WidgetRight) As String

        Dim request As HttpWebRequest
        Dim response As HttpWebResponse
        Dim stream As Stream
        Dim streamRead As StreamReader
        Dim jsonReturn As String
        Dim fieldID As String = ""


        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = (SecurityProtocolType.Tls12)

        Dim jsonData As String

        jsonData = "{" & ControlChars.CrLf
        jsonData = jsonData & "   " & ControlChars.Quote & "name" & ControlChars.Quote & ": " & ControlChars.Quote & fieldName & ControlChars.Quote & "," & ControlChars.CrLf
        jsonData = jsonData & "   " & ControlChars.Quote & "description" & ControlChars.Quote & ": " & ControlChars.Quote & SignatureDescription & ControlChars.Quote & "," & ControlChars.CrLf
        jsonData = jsonData & "   " & ControlChars.Quote & "signerId" & ControlChars.Quote & ": " & ControlChars.Quote & SignerID & ControlChars.Quote & "," & ControlChars.CrLf
        jsonData = jsonData & "   " & ControlChars.Quote & "alternateName" & ControlChars.Quote & ": " & ControlChars.Quote & alternativeName & ControlChars.Quote & "," & ControlChars.CrLf
        jsonData = jsonData & "   " & ControlChars.Quote & "required" & ControlChars.Quote & ": " & LCase(RequiredField) & "," & ControlChars.CrLf
        jsonData = jsonData & "   " & ControlChars.Quote & "readOnly" & ControlChars.Quote & ": " & LCase(ReadOnlyField) & "," & ControlChars.CrLf
        jsonData = jsonData & "   " & ControlChars.Quote & "widgets" & ControlChars.Quote & ": [" & ControlChars.CrLf
        jsonData = jsonData & "      " & "{" & ControlChars.CrLf
        jsonData = jsonData & "      " & ControlChars.Quote & "index" & ControlChars.Quote & ": " & WidgetIndex & "," & ControlChars.CrLf
        jsonData = jsonData & "      " & ControlChars.Quote & "pageNumber" & ControlChars.Quote & ": " & WidgetPageNumber & "," & ControlChars.CrLf
        jsonData = jsonData & "      " & ControlChars.Quote & "top" & ControlChars.Quote & ": " & WidgetTop & "," & ControlChars.CrLf
        jsonData = jsonData & "      " & ControlChars.Quote & "left" & ControlChars.Quote & ": " & WidgetLeft & "," & ControlChars.CrLf
        jsonData = jsonData & "      " & ControlChars.Quote & "right" & ControlChars.Quote & ": " & WidgetRight & "," & ControlChars.CrLf
        jsonData = jsonData & "      " & ControlChars.Quote & "bottom" & ControlChars.Quote & ": " & WidgetBottom & "," & ControlChars.CrLf
        jsonData = jsonData & "      " & ControlChars.Quote & "tabIndex" & ControlChars.Quote & ": " & "0" & ControlChars.CrLf
        jsonData = jsonData & "      " & "}" & ControlChars.CrLf
        jsonData = jsonData & "   " & "]" & ControlChars.CrLf & "}"


        Dim serverURL As String = serverAddress + "/rest/v7/packages/" & PackageID & "/documents/" & DocumentID & "/signaturefield?resolution=72&natural_order=true"
        Try
            request = WebRequest.Create(serverURL)
            request.ContentType = "application/json"
            request.Accept = "application/json"
            request.Method = "POST"
            request.Headers.Add("X-AUTH-TOKEN", authToken)

            Using streamWriter As New StreamWriter(request.GetRequestStream())
                streamWriter.Write(jsonData)
            End Using
            response = request.GetResponse

            If response.StatusCode = 201 Then
                stream = response.GetResponseStream
                streamRead = New StreamReader(stream)
                jsonReturn = streamRead.ReadToEnd
                streamRead.Close()
                stream.Close()

                Dim FieldDetails As SignDocSignerDetails = JsonConvert.DeserializeObject(Of SignDocSignerDetails)(jsonReturn)

                fieldID = FieldDetails.id
            End If

        Catch ex As Exception

            fieldID = ex.Message


        End Try
        Return fieldID



    End Function

    ''' <summary>
    ''' Schedule the SignDoc package for Sending, effectively sends the SignDoc Package identified by PackageID to the Signers
    ''' </summary>
    ''' <param name="serverAddress">URL address of the SignDoc Sever</param>
    ''' <param name="authToken">Authentication token from the GetAuthenticationToken function</param>
    ''' <param name="PackageID">Package ID from the CreatePackage function</param>
    ''' <returns></returns>
    Public Shared Function SchedulePackage(ByVal serverAddress, ByVal authToken, ByVal PackageID) As String

        Dim request As HttpWebRequest
        Dim response As HttpWebResponse
        Dim callReturn As String = ""

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = (SecurityProtocolType.Tls12)

        Dim serverURL As String = serverAddress + "/rest/v7/packages/" & PackageID & "/scheduler"

        Try
            request = WebRequest.Create(serverURL)
            request.ContentType = "application/json"
            request.Method = "POST"
            request.Headers.Add("X-AUTH-TOKEN", authToken)

            response = request.GetResponse

            If response.StatusCode = System.Net.HttpStatusCode.OK Then
                callReturn = response.StatusCode
            End If
        Catch ex As Exception
            callReturn = ex.Message & " Call: " & serverURL
        End Try

        Return callReturn


    End Function
End Class

Public Class SignDocPackage
    Public Property name As String
    Public Property type As String

End Class

Public Class SignDocPackageDetails
    Public Property id As String
    Public Property url As String

End Class

Public Class SignerDetails

    Public Property email As String
    Public Property name As String
    Public Property role As String


End Class

Public Class SignDocSignerDetails
    Public Property id As String
    Public Property messages As String
    Public Property url As String

End Class

Public Class DocumentEntry

    ' Document Content as BASE64
    Public Property content As String
    Public Property description As String
    Public Property documentMessage As String
    Public Property fileName As String
    Public Property format As String
    Public Property name As String


End Class

Public Class DocumentReturn
    Public Property description As String
    Public Property documentMessage As String
    Public Property fileName As String
    Public Property id As String
    Public Property name As String
    Public Property order As String
    Public Property thumbnai As String

    Public Property url As String


End Class
