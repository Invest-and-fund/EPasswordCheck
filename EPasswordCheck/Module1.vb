Imports System.Configuration

Module Module1

    Public SystemLocation As Integer = 0  ' 0 = Live  1=local

    Public Enum EquPassType
        ID
        Bank
    End Enum

    Public Property BackupFileName As String
    Public Property BankResponse As String
    Public ReadOnly Property DbaseLoc As String = ConfigurationManager.AppSettings("DatabaseLocation")
    Public ReadOnly Property DbasePwd As String = ConfigurationManager.AppSettings("DatabasePassword")
    Public ReadOnly Property DbaseUsr As String = ConfigurationManager.AppSettings("DatabaseUser")
    Public ReadOnly Property FirebirdLoc As String = ConfigurationManager.AppSettings("FireBirdInstallLocation")
    Public Property IDresponse As String
    Public ReadOnly Property LBorDocsFolder As String = ConfigurationManager.AppSettings("LFolderPath") & ConfigurationManager.AppSettings("BorrowerDocsFolder")
    Public ReadOnly Property LConnStr As String = ConfigurationManager.ConnectionStrings("LFBConnectionString").ConnectionString
    Public ReadOnly Property LDbaseNam As String = ConfigurationManager.AppSettings("LDatabaseName")
    Public ReadOnly Property LLoanImgsFolder As String = ConfigurationManager.AppSettings("LFolderPath") & ConfigurationManager.AppSettings("LoanImgsFolder")
    Public ReadOnly Property PortNumber As Integer = CInt(ConfigurationManager.AppSettings("PortNum"))
    Public Property ReturnString As String
    Public ReadOnly Property SBorDocsFolder As String = ConfigurationManager.AppSettings("SFolderPath") & ConfigurationManager.AppSettings("BorrowerDocsFolder")
    Public ReadOnly Property SConnStr As String = ConfigurationManager.ConnectionStrings("SFBConnectionString").ConnectionString
    Public ReadOnly Property SDbaseNam As String = ConfigurationManager.AppSettings("SDatabaseName")
    Public ReadOnly Property SLoanImgsFolder As String = ConfigurationManager.AppSettings("SFolderPath") & ConfigurationManager.AppSettings("LoanImgsFolder")
    Public ReadOnly Property TestMode As Boolean = ConfigurationManager.AppSettings("TestMode") = 0
    Function DoBankCheck() As Boolean
        DoBankCheck = False

        Dim ReqTemplate As New Xml.XmlDocument
        ReqTemplate.LoadXml("<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v1=""http://ewsconsumer.services.uk.equifax.com/schema/v1/creditsearch/creditsearchrequest"">" &
                            "<soapenv:Header><wsse:Security soapenv:mustUnderstand=""1"" xmlns:wsse=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"">" &
                            "<wsse:UsernameToken wsu:Id=""UsernameToken-2"" xmlns:wsu=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"">" &
                            "<wsse:Username>INVESTANDFUND@INVESTXML2</wsse:Username><wsse:Password Type=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText"">" &
                            "</wsse:Password></wsse:UsernameToken></wsse:Security></soapenv:Header><soapenv:Body>" &
                            "<ns10:verifyIdentityCheckRequest xmlns:ns10=""http://ewsconsumer.services.uk.equifax.com/schema/v1/identityverification/verifyidentitycheckrequest""><clientRef>STBCA1 Request</clientRef>" &
                            "<soleSearch><matchCriteria><associate>notRequired</associate><attributable>notRequired</attributable><family>notRequired</family><potentialAssociate>notRequired</potentialAssociate>" &
                            "<subject>required</subject></matchCriteria><requestedData><scoreAndCharacteristicRequests><employSameCompanyInsight>true</employSameCompanyInsight><characteristicRequests><index>1</index>" &
                            "<taggedCharacteristics>true</taggedCharacteristics></characteristicRequests></scoreAndCharacteristicRequests></requestedData><primary><bankingInfo><account><accountNumber></accountNumber>" &
                            "<sortCode></sortCode></account></bankingInfo><dob></dob><name><middleName/><surname></surname><title/><forename></forename></name><currentAddress><address><county></county>" &
                            "<district></district><number></number><postcode></postcode><postTown></postTown><street1></street1></address></currentAddress></primary></soleSearch></ns10:verifyIdentityCheckRequest>" &
                            "</soapenv:Body></soapenv:Envelope>")
        Dim nsmgr As New Xml.XmlNamespaceManager(ReqTemplate.NameTable)
        nsmgr.AddNamespace("soapenv", "http://schemas.xmlsoap.org/soap/envelope/")
        nsmgr.AddNamespace("wsse", "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd")
        nsmgr.AddNamespace("wsu", "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd")
        nsmgr.AddNamespace("ns2", "http://ewsconsumer.services.uk.equifax.com/schema/v1/creditsearch/creditsearchrequest")
        nsmgr.AddNamespace("ns10", "http://ewsconsumer.services.uk.equifax.com/schema/v1/identityverification/verifyidentitycheckrequest")

        ReqTemplate.SelectSingleNode("soapenv:Envelope/soapenv:Header/wsse:Security/wsse:UsernameToken/wsse:Password", nsmgr).InnerText = GetPassword(EquPassType.Bank)
        ReqTemplate.SelectSingleNode("soapenv:Envelope/soapenv:Body/ns10:verifyIdentityCheckRequest/soleSearch/primary/bankingInfo/account/accountNumber", nsmgr).InnerText = "00000000"
        ReqTemplate.SelectSingleNode("soapenv:Envelope/soapenv:Body/ns10:verifyIdentityCheckRequest/soleSearch/primary/bankingInfo/account/sortCode", nsmgr).InnerText = "00-00-00"
        ReqTemplate.SelectSingleNode("soapenv:Envelope/soapenv:Body/ns10:verifyIdentityCheckRequest/soleSearch/primary/dob", nsmgr).InnerText = "2019-01-01"
        ReqTemplate.SelectSingleNode("soapenv:Envelope/soapenv:Body/ns10:verifyIdentityCheckRequest/soleSearch/primary/name/surname", nsmgr).InnerText = "McTestFace"
        ReqTemplate.SelectSingleNode("soapenv:Envelope/soapenv:Body/ns10:verifyIdentityCheckRequest/soleSearch/primary/name/forename", nsmgr).InnerText = "Testy"
        ReqTemplate.SelectSingleNode("soapenv:Envelope/soapenv:Body/ns10:verifyIdentityCheckRequest/soleSearch/primary/currentAddress/address/number", nsmgr).InnerText = "123"
        ReqTemplate.SelectSingleNode("soapenv:Envelope/soapenv:Body/ns10:verifyIdentityCheckRequest/soleSearch/primary/currentAddress/address/postcode", nsmgr).InnerText = "HZ109FH"
        ReqTemplate.SelectSingleNode("soapenv:Envelope/soapenv:Body/ns10:verifyIdentityCheckRequest/soleSearch/primary/currentAddress/address/postTown", nsmgr).InnerText = "Fake City"
        ReqTemplate.SelectSingleNode("soapenv:Envelope/soapenv:Body/ns10:verifyIdentityCheckRequest/soleSearch/primary/currentAddress/address/street1", nsmgr).InnerText = "Fake Street"

        Dim Client As New Net.WebClient
        Client.Headers.Add("Content-Type", "text/xml;charset=utf-8")
        Client.Headers.Add("SOAPAction", """https://services.uk.equifax.com/xmlii/EWSConsumerService-v1_x/consumerService.wsdl""")
        Dim sURL As String = "https://services.uk.equifax.com/xmlii/EWSConsumerService-v1_x/consumerService"

        Try
            BankResponse = Client.UploadString(sURL, ReqTemplate.OuterXml)
        Catch ex As Net.WebException
            Dim resstream As IO.Stream = ex.Response.GetResponseStream()
            Dim sr As New IO.StreamReader(resstream)
            BankResponse = sr.ReadToEnd()
        End Try

        Dim ExpectedResponse As String = "<\?xml version=" & "'" & "1\.0" & "'" & " encoding=" & "'" & "UTF-8" & "'" & "\?><soapenv:Envelope xmlns:soapenv=""http:\/\/schemas\.xmlsoap\.org\/soap\/envelope\/""><soapenv:Header><sec:PasswordExpiryInformation xmlns:sec=""http:\/\/ewsconsumer\.services\.uk\.equifax\.com\/schema\/v1\/security""><PasswordSecurity xmlns=""http:\/\/services\.uk\.equifax\.com\/schema\/v1\/security""><mustChangePassword>(.*)<\/mustChangePassword><passwordValidityPeriod>(.*)<\/passwordValidityPeriod><latestDateOfPasswordChange>(.*)<\/latestDateOfPasswordChange><lastPasswordChange>(.*)<\/lastPasswordChange><\/PasswordSecurity><\/sec:PasswordExpiryInformation><\/soapenv:Header><soapenv:Body><ns2:verifyIdentityCheckResponse xmlns:ns2=""http:\/\/ewsconsumer\.services\.uk\.equifax\.com\/schema\/v1\/identityverification\/verifyidentitycheckresponse""><clientRef>STBCA1 REQUEST<\/clientRef><soleSearch><primary><suppliedAddressData><addressMatchStatus>(.*)<\/addressMatchStatus><index>(.*)<\/index><\/suppliedAddressData><\/primary><\/soleSearch><\/ns2:verifyIdentityCheckResponse><\/soapenv:Body><\/soapenv:Envelope>"

        If System.Text.RegularExpressions.Regex.IsMatch(BankResponse, ExpectedResponse) Then
            DoBankCheck = True
        End If

    End Function

    Function DoIDCheck() As Boolean
        DoIDCheck = False
        Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12

        Dim bodyString As String = "group=018750&user=SKENNY&password=" & GetPassword(EquPassType.ID) &
            "&request_type=ADDRESS&request_subtype=ADDRESSMATCH&version=5" &
            "&forename=Testy" &
            "&surname=McTestFace" &
            "&house_name=123" &
            "&postcode=HZ109FH" &
            "&date_of_birth=2017-01-01"

        Dim requestString As String = "https://www.uk.equifax.com/servlet/CobaltXML"


        Dim request = Net.WebRequest.Create(requestString)
        request.Method = "POST"
        request.ContentType = "application/x-www-form-urlencoded"
        Dim encoding As New System.Text.ASCIIEncoding()
        Dim _byte() As Byte = encoding.GetBytes(bodyString)
        request.ContentLength = _byte.Length
        Dim requestStream As IO.Stream = request.GetRequestStream
        requestStream.Write(_byte, 0, _byte.Length)
        requestStream.Close()

        Dim response = request.GetResponse()
        Dim reader As New IO.StreamReader(response.GetResponseStream)
        IDresponse = Trim(reader.ReadToEnd)

        Dim ExpectedResponse As String = "<\?xml version=""1\.0"" encoding=""UTF-8""\?><response schema_version=""2\.2""><response_header session_token=""(.*)"">" &
            "<client_reference>(.*)<\/client_reference><\/response_header><service_response id=""1"" success_flag=""1""><address_matching_response><match_address_response>" &
            "<return_address county="""" district="""" house_name="""" house_no="""" match_status=""0"" post_town="""" postcode="""" ptc_abs_code="""" request_id=""0"" street_1="""" street_2="""" surname=""""\/>" &
            "<\/match_address_response><\/address_matching_response><\/service_response><image_url>(.*)<\/image_url><\/response>"

        If System.Text.RegularExpressions.Regex.IsMatch(IDresponse, ExpectedResponse) Then
            DoIDCheck = True
        End If

    End Function

    Function GetPassword(ByVal Type As EquPassType) As String

        GetPassword = Nothing

        Select Case Type
            Case EquPassType.ID
                Dim Cmd = New FirebirdSql.Data.FirebirdClient.FbCommand With {
                    .Connection = New FirebirdSql.Data.FirebirdClient.FbConnection(ConfigurationManager.ConnectionStrings("LFBConnectionString").ToString),
                    .CommandType = CommandType.Text,
                    .CommandText = "SELECT first 1 a.SPASSWORD FROM EQUIFAX_PW a order by a.EQUIFAX_PW_ID desc"
                }
                Cmd.Connection.Open()
                Dim Reader As FirebirdSql.Data.FirebirdClient.FbDataReader
                Try
                    Reader = Cmd.ExecuteReader()
                    If Reader.Read() Then
                        GetPassword = Trim(Reader.Item(0))
                    End If
                Catch ex As Exception

                End Try
                Cmd.Connection.Close()
                Cmd.Connection = Nothing
                Cmd = Nothing
            Case EquPassType.Bank
                Dim Cmd = New FirebirdSql.Data.FirebirdClient.FbCommand With {
                    .Connection = New FirebirdSql.Data.FirebirdClient.FbConnection(ConfigurationManager.ConnectionStrings("LFBConnectionString").ToString),
                    .CommandType = CommandType.Text,
                    .CommandText = "SELECT first 1 a.BPASSWORD FROM EQUIFAX_PW a order by a.EQUIFAX_PW_ID desc"
                }
                Cmd.Connection.Open()
                Dim Reader As FirebirdSql.Data.FirebirdClient.FbDataReader
                Try
                    Reader = Cmd.ExecuteReader()
                    If Reader.Read() Then
                        GetPassword = Trim(Reader.Item(0))
                    End If
                Catch ex As Exception

                End Try
                Cmd.Connection.Close()
                Cmd.Connection = Nothing
                Cmd = Nothing
        End Select

    End Function

    Sub Main()
        Dim sPW As String = ConfigurationManager.AppSettings("EPCPW")
        Dim sUSR As String = ConfigurationManager.AppSettings("EPCUSR")
        Dim mail As New System.Net.Mail.MailMessage With {
            .From = New Net.Mail.MailAddress(sUSR),
            .Subject = "EquiFax Password Check Response"
        }
        mail.To.Add("web@investandfund.com")

        Dim bodyText As String = "ID password check:<br>" &
            IIf(DoIDCheck(), "Passed", "Failed") & "<br><br>" &
            "ID check response string:<br>" &
            System.Net.WebUtility.HtmlEncode(IDresponse) & "<br><br><br><br><br>" &
            "Bank password check:<br>" &
            IIf(DoBankCheck(), "Passed", "Failed") & "<br><br>" &
            "Bank check response string:<br>" &
            System.Net.WebUtility.HtmlEncode(BankResponse)

        mail.Body = bodyText
        mail.IsBodyHtml = True

        Dim smtp As New System.Net.Mail.SmtpClient() With {
            .Host = "smtp.office365.com",
            .Credentials = New System.Net.NetworkCredential(sUSR, sPW),
            .EnableSsl = True,
            .Port = 587
        }

        smtp.Send(mail)

    End Sub

End Module