Option Explicit
Dim rowCount     As Long
Dim colCount     As Long
Dim valRow()     As Variant
Dim ScriptEngine As Object

'*************************************************************************************************************
Public Function Keys(ByVal jsonOb As Object) As String()
'*************************************************************************************************************   
    Dim KeysOb As Object
    Dim Key As Variant
    Dim KeyList() As String
    Dim i As Integer
    
    Set KeysOb = ScriptEngine.Run("getKeys", jsonOb)
    ReDim KeyList(ScriptEngine.Run("getProperty", KeysOb, "length") - 1)
    
    For Each Key In KeysOb
        KeyList(i) = Key
        i = i + 1
    Next
    
    Keys = KeyList
End Function

'*************************************************************************************************************
Public Function JSON2Array(ByVal strResponse As String) As Variant()
'*************************************************************************************************************
    Dim jsonOb As Object
    Dim Arr() As Variant
    
    
    colCount = 0
    rowCount = 0
    ReDim valRow(0)
    
    Set ScriptEngine = CreateObject("MSScriptcontrol.scriptControl")
    ScriptEngine.Language = "JScript"
    ScriptEngine.AddCode "function getProperty(jsonObj, property) { return jsonObj[property]; } "
    ScriptEngine.AddCode "function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } "
    
    Set jsonOb = ScriptEngine.Eval("(" + strResponse + ")")
    
   GetRecData jsonOb, "", JSON2Array
   
End Function

'*************************************************************************************************************
Sub GetRecData(ByVal ParentOb As Object, ByVal property As String, finalArray() As Variant)
'*************************************************************************************************************   
   Dim KeyName, KeyVal As Variant
   Dim i, j As Integer
   Dim ChildOb As Object
   Dim headerRow() As Variant
   
    If property <> "" Then
        KeyVal = ScriptEngine.Run("getProperty", ParentOb, property)
        If InStr(KeyVal, "[object Object]") > 0 Then
            Set ChildOb = ScriptEngine.Run("getProperty", ParentOb, property)
        Else
            Return
        End If
    Else
        Set ChildOb = ParentOb
    End If
    
    If Not ChildOb Is Nothing Then
        colCount = colCount + 1
        For Each KeyName In Keys(ChildOb)
            
            ReDim Preserve valRow(colCount - 1)
            valRow(UBound(valRow)) = KeyName
            
            KeyVal = ScriptEngine.Run("getProperty", ChildOb, KeyName)
            If InStr(KeyVal, "[object Object]") > 0 Then
                Call GetRecData(ChildOb, KeyName, finalArray)
            Else
                    
                ReDim Preserve valRow(colCount)
                valRow(UBound(valRow)) = KeyVal
                rowCount = rowCount + 1
                ReDim Preserve finalArray(rowCount - 1)
                finalArray(UBound(finalArray)) = valRow()
            End If
        Next KeyName
        ReDim Preserve valRow(colCount - 1)
        colCount = colCount - 1
   
   End If
End Sub

'*************************************************************************************************************
Sub DemoJsonToArray()
'*************************************************************************************************************
Dim strJson As String
Dim valTab() As Variant
strJson = "{""Bihar"": {""district"": {""Katihar"":{""male"": 10000000,""female"": 800000,""age-group"": {""0-17"": 1000000,""18-59"": 600000,""60-120"": 200000}},""Darbhanga"": {""male"": 8000000,""female"": 700000,""age-group"":{""0-17"": 600000,""18-59"": 800000,""60-120"": 100000}}}},""Maharashtra"": {""district"": {""Ahmednagar"": {""male"": 6000000,""female"": 400000,""age-group"": {""0-17"": 500000,""18-59"": 400000,""60-120"": 100000}},""Mumbai"":{""male"": 80000000,""female"": 7500000,""age-group"": {""0-17"": 5000000,""18-59"": 8000000,""60-120"": 2500000}}}}}"
valTab() = JSON2Array(strJson)
moveToSheet valTab, 1
End Sub

'*************************************************************************************************************
Sub moveToSheet(valTab() As Variant, sheetNum As Integer)
'*************************************************************************************************************
    Dim i, j
    
    Sheets(sheetNum).Cells.Clear
    
    For i = 0 To UBound(valTab)
        For j = 0 To UBound(valTab(i))
            Sheets(sheetNum).Cells(i + 1, j + 1).Value = valTab(i)(j)
        Next j
    Next i
    Sheets(sheetNum).Select
End Sub
'*************************************************************************************************************


