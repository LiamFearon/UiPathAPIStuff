Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Configuration
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports Nancy.Json
Imports System
Imports System.Diagnostics

Module OrchAPI

    Sub Liam()
        Dim strOrchURL As String = "https://cloud.uipath.com/robiqubrergn/"
        Dim strClientID As String = "8DEv1AMNXczW3y4U15LL3jYf62jK93n5"
        Dim strRefreshToken As String = "D6FujDr96mxqZKR8vXfwOt24ONQMp6Pr7Ht59jkRSoKhV"
        Dim strTenantName As String = "DefaultTenant"
        Dim strBearerToken As String = GetToken(strClientID, strRefreshToken, strTenantName)

        Console.WriteLine(GetQueueItems(strOrchURL, strTenantName, strBearerToken))
        Console.ReadKey()



    End Sub

    ''' <summary>
    ''' Taking a users API credentials (Found in orchestrator), a bearer token is recovered for use in further API calls that require authentication as a String
    ''' </summary>
    ''' <param name="ClientID"></param>
    ''' <param name="RefreshToken"></param>
    ''' <param name="TenantName"></param>
    ''' <returns></returns>
    Public Function GetToken(ClientID As String, RefreshToken As String, TenantName As String) As String
        Dim myReq As Net.HttpWebRequest = HttpWebRequest.Create("https://account.uipath.com/oauth/token") ' Generic request address for users
        myReq.Method = "POST"                                                                             ' Call type
        myReq.ContentType = "application/json"                                                            ' Response format
        myReq.Timeout = 5000                                                                              ' Timeout value
        myReq.Headers.Add("X-UIPATH-TenantName", TenantName)                                              ' Required header for call

        'Building body of call
        Dim PostString As String
        PostString = "{
                ""grant_type"": ""refresh_token"",
                ""client_id"": " + """" + ClientID + """" + ",
                ""refresh_token"": " + """" + RefreshToken + """" + "
        }"

        Dim byteArray As Byte() = Encoding.UTF8.GetBytes(PostString)
        myReq.ContentLength = byteArray.Length
        Dim dataStream As Stream

        dataStream = myReq.GetRequestStream()
        dataStream.Write(byteArray, 0, byteArray.Length)
        dataStream.Close()

        Try
            Debug.WriteLine("Sending API request for bearer token")
            Dim response As HttpWebResponse = CType(myReq.GetResponse(), HttpWebResponse)
            Dim receiveStream As Stream = response.GetResponseStream()
            Dim readStream As New StreamReader(receiveStream, Encoding.UTF8)
            Dim res = readStream.ReadToEnd()

            response.Close()
            readStream.Close()

            Debug.WriteLine("Request sent")
            Dim j As Object = New JavaScriptSerializer().Deserialize(Of Object)(res)

            Return j("access_token")
        Catch ex As Exception
            Debug.WriteLine("GetToken failed: " + ex.Message)
            Throw
        End Try
    End Function


    ''' <summary>
    ''' Fetches all available queue items from 
    ''' </summary>
    ''' <param name="ClientID"></param>
    ''' <param name="RefreshToken"></param>
    ''' <param name="TenantName"></param>
    ''' <returns></returns>
    Public Function GetQueueItems(OrchURL As String, TenantName As String, BearerToken As String) As String
        Dim reqBuilder As String = OrchURL + TenantName + "/odata/QueueItems()"
        Dim myReq As Net.HttpWebRequest = HttpWebRequest.Create(reqBuilder)
        myReq.Method = "GET"
        myReq.ContentType = "application/json"
        myReq.Timeout = 10000


        myReq.Headers.Add("Authorization", "Bearer " & BearerToken)
        myReq.Headers.Add("X-UIPATH-OrganizationUnitId", "1119257")
        myReq.Headers.Add("Cookie", "UiPathBrowserId=dffa2993-9103-41c5-90ed-b5ad46003700")

        Dim response As HttpWebResponse = CType(myReq.GetResponse(), HttpWebResponse)

        Console.WriteLine("Content length is {0}", response.ContentLength)
        Console.WriteLine("Content type is {0}", response.ContentType)

        Dim receiveStream As Stream = response.GetResponseStream()

        ' Pipes the stream to a higher level stream reader with the required encoding format. 
        Dim readStream As New StreamReader(receiveStream, Encoding.UTF8)

        Console.WriteLine("Response stream received.")
        Return readStream.ReadToEnd()

        response.Close()
        readStream.Close()

    End Function


End Module