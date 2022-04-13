Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Imports System

Module Module1

    Sub Main()
        OrchAPI.Liam()

    End Sub


    Sub AddQueueItem(ByVal strBearerToken As String)


        Dim myReq As Net.HttpWebRequest = HttpWebRequest.Create("https://cloud.uipath.com/TomsTestOrch/DefaultTenant/odata/Queues/UiPathODataSvc.AddQueueItem")
        myReq.Method = "POST"
        myReq.ContentType = "application/json"
        myReq.Timeout = 100000

        myReq.Headers.Add("Authorization", "Bearer " & strBearerToken)
        myReq.Headers.Add("X-UIPATH-TenantName", "DefaultTenant")
        myReq.Headers.Add("X-UIPATH-FolderPath", "Toms Folder")
        myReq.Headers.Add("Cookie", "UiPathBrowserId=f75f2b58-4670-4525-bbc3-901dff64d904")

        Dim PostString As String '= JsonConvert.SerializeObject(NewData)

        PostString = "{
	""itemData"": {
		""Priority"": ""High"",
		""DeferDate"": ""2023-03-21T13:42:27.654Z"",
		""DueDate"": ""2022-03-25T13:42:27.654Z"",
		""Name"": ""CAD_ExportQueue"",
		""SpecificContent"": {
			""Email"": ""barry@uipath.com"", 
			""Name"": ""O'Brian""
		}
	}
}"
        Dim byteArray As Byte() = Encoding.UTF8.GetBytes(PostString)
        myReq.ContentLength = byteArray.Length
        Dim dataStream As Stream

        dataStream = myReq.GetRequestStream()
        dataStream.Write(byteArray, 0, byteArray.Length)
        dataStream.Close()

        Dim myResp As HttpWebResponse = myReq.GetResponse()




    End Sub

    Function GetQueueItem(ByVal strBearerToken As String) As String


        Dim myReq As Net.HttpWebRequest = HttpWebRequest.Create("https://cloud.uipath.com/TomsTestOrch/DefaultTenant/odata/QueueItems?$filter=Status eq 'New'&$top=10")
        myReq.Method = "GET"
        myReq.ContentType = "application/json"
        myReq.Timeout = 10000

        myReq.Headers.Add("Authorization", "Bearer " & strBearerToken)

        Dim response As HttpWebResponse = CType(myReq.GetResponse(), HttpWebResponse)

        Console.WriteLine("Content length is {0}", response.ContentLength)
        Console.WriteLine("Content type is {0}", response.ContentType)

        Dim receiveStream As Stream = response.GetResponseStream()

        ' Pipes the stream to a higher level stream reader with the required encoding format. 
        Dim readStream As New StreamReader(receiveStream, Encoding.UTF8)

        Console.WriteLine("Response stream received.")
        GetQueueItem = readStream.ReadToEnd()

        response.Close()
        readStream.Close()

    End Function


    Sub RegexDataSort(ByVal strInput As String, ByVal strPattern As String)

        Console.WriteLine("Starting Regex Match..")

        Dim matches As MatchCollection = Regex.Matches(strInput, strPattern)
        Dim intMatchCount As Int16 = 0

        Console.WriteLine("Matches Found: " & matches.Count)


        For Each match As Match In matches
            Console.WriteLine()
            Console.WriteLine("Case Number:    {0}", intMatchCount + 1)
            Console.WriteLine("Queue ID:       {0}", match.Groups(1).Value)
            Console.WriteLine("Status:         {0}", match.Groups(2).Value)
            Console.WriteLine("Key:            {0}", match.Groups(3).Value)
            Console.WriteLine("Exception Type: {0}", match.Groups(4).Value)
            Console.WriteLine("Email:          {0}", match.Groups(5).Value)
            Console.WriteLine("Name:           {0}", match.Groups(6).Value)
            Console.WriteLine()

            intMatchCount += 1

        Next


        Console.WriteLine()

    End Sub



    Public Class Container
        Public Venue As JSON_result
    End Class

    Public Class JSON_result
        Public QueueDefinitionId As Integer
        Public Status As String

    End Class


    Sub DeserialiseJSON(ByVal jstr As String)

        'Dim obj = JsonConvert.DeserializeObject(Of Container)(jstr)

        Console.WriteLine()

    End Sub


End Module
