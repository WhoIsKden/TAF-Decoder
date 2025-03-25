# TAF-Decoder
Sub GetTAFData()
    Dim url As String
    Dim http As Object
    Dim json As Object
    Dim tafForecast As Object
    Dim entry As Object
    Dim timePeriod As String
    Dim forecast As String
    
    ' API URL for KDLF TAF data, valid for 30 hours
    url = "https://aviationweather.gov/api/data/taf?ids=KDLF&hours=30&sep=true"
    
    ' Create the HTTP request object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Send the GET request to fetch the TAF data
    http.Open "GET", url, False
    http.Send
    
    ' Parse the JSON response
    Set json = JsonConverter.ParseJson(http.responseText) ' Ensure you have the JSON parser
    
    ' Access the 'data' array in the response JSON
    Set tafForecast = json("data")
    
    ' Loop through each TAF entry and extract details
    For Each entry In tafForecast
        ' Get the time period and forecast
        timePeriod = entry("time")
        forecast = entry("forecast")
        
        ' Output the results to the Immediate Window (Ctrl + G to view)
        Debug.Print "Time Period: " & timePeriod
        Debug.Print "Forecast: " & forecast
        Debug.Print ""
    Next entry
End Sub
