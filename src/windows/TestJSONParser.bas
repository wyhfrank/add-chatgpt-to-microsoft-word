' Add a reference to "Microsoft Scripting Runtime" and "Microsoft Script Control x.x" in the VBA editor

Sub ParseJSON()
    ' Create a new instance of the ScriptControl
    Dim sc As New ScriptControl
    sc.Language = "JScript"
    
    ' Load the JSON conversion functions
    sc.AddCode "function parseJSON(json) { return JSON.parse(json); }"
    sc.AddCode "function stringifyJSON(obj) { return JSON.stringify(obj); }"
    
    ' JSON data
    Dim jsonData As String
    jsonData = "{""name"": ""John Smith"", ""age"": 30, ""city"": ""New York""}"
    
    ' Parse JSON
    Dim parsedData As Object
    Set parsedData = sc.Run("parseJSON", jsonData)
    
    ' Access the parsed data
    Dim name As String
    Dim age As Integer
    Dim city As String
    
    name = parsedData("name")
    age = parsedData("age")
    city = parsedData("city")
    
    ' Display the parsed data
    MsgBox "Name: " & name & vbNewLine & _
           "Age: " & age & vbNewLine & _
           "City: " & city
End Sub
