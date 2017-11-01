Attribute VB_Name = "JSON"

' VBA JSON parser, Backus–Naur form JSON parser based on RegEx v1.6.01
' Copyright (C) 2015-2017 omegastripes
' omegastripes@yandex.ru
' https://github.com/omegastripes/VBA-JSON-parser
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.

Option Explicit

Private sBuffer As String
Private oTokens As Object
Private oRegEx As Object
Private bMatch As Boolean
Private oChunks As Object
Private oHeader As Object
Private aData() As Variant
Private i As Long
Private sDelim As String
    
Sub Parse(ByVal sSample As String, vJSON As Variant, sState As String)
    
    ' Input:
    ' sSample - source JSON string
    ' Output:
    ' vJson - created object or array to be returned as result
    ' sState - string Object|Array|Error depending on result
    
    sBuffer = sSample
    Set oTokens = CreateObject("Scripting.Dictionary")
    Set oRegEx = CreateObject("VBScript.RegExp")
    With oRegEx ' Patterns based on specification http://www.json.org/
        .Global = True
        .MultiLine = True
        .IgnoreCase = True ' Unspecified True, False, Null accepted
        .Pattern = "(?:'[^']*'|""(?:\\""|[^""])*"")(?=\s*[,\:\]\}])" ' Double-quoted string, unspecified quoted string
        Tokenize "s"
        .Pattern = "[+-]?(?:\d+\.\d*|\.\d+|\d+)(?:e[+-]?\d+)?(?=\s*[,\]\}])" ' Number, E notation number
        Tokenize "d"
        .Pattern = "\b(?:true|false|null)(?=\s*[,\]\}])" ' Constants true, false, null
        Tokenize "c"
        .Pattern = "\b[A-Za-z_]\w*(?=\s*\:)" ' Unspecified non-double-quoted property name accepted
        Tokenize "n"
        .Pattern = "\s{2,}"
        sBuffer = .Replace(sBuffer, "") ' Remove unnecessary spaces
        .MultiLine = False
        Do
            bMatch = False
            .Pattern = "<\d+(?:[sn])>\:<\d+[codas]>" ' Object property structure
            Tokenize "p"
            .Pattern = "\{(?:<\d+p>(?:,<\d+p>)*)?\}" ' Object structure
            Tokenize "o"
            .Pattern = "\[(?:<\d+[codas]>(?:,<\d+[codas]>)*)?\]" ' Array structure
            Tokenize "a"
        Loop While bMatch
        .Pattern = "^<\d+[oa]>$" ' Top level object structure, unspecified array accepted
        If .Test(sBuffer) And oTokens.Exists(sBuffer) Then
            sDelim = Mid(1 / 2, 2, 1)
            Retrieve sBuffer, vJSON
            sState = IIf(IsObject(vJSON), "Object", "Array")
        Else
            vJSON = Null
            sState = "Error"
        End If
    End With
    Set oTokens = Nothing
    Set oRegEx = Nothing
    
End Sub

Private Sub Tokenize(sType)
    
    Dim aContent() As String
    Dim lCopyIndex As Long
    Dim i As Long
    Dim sKey As String
    
    With oRegEx.Execute(sBuffer)
        If .Count = 0 Then Exit Sub
        ReDim aContent(0 To .Count - 1)
        lCopyIndex = 1
        For i = 0 To .Count - 1
            With .Item(i)
                sKey = "<" & oTokens.Count & sType & ">"
                oTokens(sKey) = .Value
                aContent(i) = Mid(sBuffer, lCopyIndex, .FirstIndex - lCopyIndex + 1) & sKey
                lCopyIndex = .FirstIndex + .Length + 1
            End With
        Next
    End With
    sBuffer = Join(aContent, "") & Mid(sBuffer, lCopyIndex, Len(sBuffer) - lCopyIndex + 1)
    bMatch = True
    
End Sub

Private Sub Retrieve(sTokenKey, vTransfer)
    
    Dim sTokenValue As String
    Dim sName As String
    Dim vValue As Variant
    Dim aTokens() As String
    Dim i As Long
    
    sTokenValue = oTokens(sTokenKey)
    With oRegEx
        .Global = True
        Select Case Left(Right(sTokenKey, 2), 1)
            Case "o"
                Set vTransfer = CreateObject("Scripting.Dictionary")
                aTokens = Split(sTokenValue, "<")
                For i = 1 To UBound(aTokens)
                    Retrieve "<" & Split(aTokens(i), ">", 2)(0) & ">", vTransfer
                Next
            Case "p"
                aTokens = Split(sTokenValue, "<", 4)
                Retrieve "<" & Split(aTokens(1), ">", 2)(0) & ">", sName
                Retrieve "<" & Split(aTokens(2), ">", 2)(0) & ">", vValue
                If IsObject(vValue) Then
                    Set vTransfer(sName) = vValue
                Else
                    vTransfer(sName) = vValue
                End If
            Case "a"
                aTokens = Split(sTokenValue, "<")
                If UBound(aTokens) = 0 Then
                    vTransfer = Array()
                Else
                    ReDim vTransfer(0 To UBound(aTokens) - 1)
                    For i = 1 To UBound(aTokens)
                        Retrieve "<" & Split(aTokens(i), ">", 2)(0) & ">", vValue
                        If IsObject(vValue) Then
                            Set vTransfer(i - 1) = vValue
                        Else
                            vTransfer(i - 1) = vValue
                        End If
                    Next
                End If
            Case "n"
                vTransfer = sTokenValue
            Case "s"
                vTransfer = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace( _
                    Mid(sTokenValue, 2, Len(sTokenValue) - 2), _
                    "\""", """"), _
                    "\\", "\"), _
                    "\/", "/"), _
                    "\b", Chr(8)), _
                    "\f", Chr(12)), _
                    "\n", vbLf), _
                    "\r", vbCr), _
                    "\t", vbTab)
                .Global = False
                .Pattern = "\\u[0-9a-fA-F]{4}"
                Do While .Test(vTransfer)
                    vTransfer = .Replace(vTransfer, ChrW(("&H" & Right(.Execute(vTransfer)(0).Value, 4)) * 1))
                Loop
            Case "d"
                vTransfer = CDbl(Replace(sTokenValue, ".", sDelim))
            Case "c"
                Select Case LCase(sTokenValue)
                    Case "true"
                        vTransfer = True
                    Case "false"
                        vTransfer = False
                    Case "null"
                        vTransfer = Null
                End Select
        End Select
    End With
    
End Sub

Function Serialize(vJSON As Variant) As String
    
    Set oChunks = CreateObject("Scripting.Dictionary")
    SerializeElement vJSON, ""
    Serialize = Join(oChunks.Items(), "")
    Set oChunks = Nothing
    
End Function

Private Sub SerializeElement(vElement As Variant, ByVal sIndent As String)
    
    Dim aKeys() As Variant
    Dim i As Long
    
    With oChunks
        Select Case VarType(vElement)
            Case vbObject
                If vElement.Count = 0 Then
                    .Item(.Count) = "{}"
                Else
                    .Item(.Count) = "{" & vbCrLf
                    aKeys = vElement.Keys
                    For i = 0 To UBound(aKeys)
                        .Item(.Count) = sIndent & vbTab & """" & aKeys(i) & """" & ": "
                        SerializeElement vElement(aKeys(i)), sIndent & vbTab
                        If Not (i = UBound(aKeys)) Then .Item(.Count) = ","
                        .Item(.Count) = vbCrLf
                    Next
                    .Item(.Count) = sIndent & "}"
                End If
            Case Is >= vbArray
                If UBound(vElement) = -1 Then
                    .Item(.Count) = "[]"
                Else
                    .Item(.Count) = "[" & vbCrLf
                    For i = 0 To UBound(vElement)
                        .Item(.Count) = sIndent & vbTab
                        SerializeElement vElement(i), sIndent & vbTab
                        If Not (i = UBound(vElement)) Then .Item(.Count) = "," 'sResult = sResult & ","
                        .Item(.Count) = vbCrLf
                    Next
                    .Item(.Count) = sIndent & "]"
                End If
            Case vbInteger, vbLong
                .Item(.Count) = vElement
            Case vbSingle, vbDouble
                .Item(.Count) = Replace(vElement, ",", ".")
            Case vbNull
                .Item(.Count) = "null"
            Case vbBoolean
                .Item(.Count) = IIf(vElement, "true", "false")
            Case Else
                .Item(.Count) = """" & _
                    Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(vElement, _
                        "\", "\\"), _
                        """", "\"""), _
                        "/", "\/"), _
                        Chr(8), "\b"), _
                        Chr(12), "\f"), _
                        vbLf, "\n"), _
                        vbCr, "\r"), _
                        vbTab, "\t") & _
                    """"
        End Select
    End With
    
End Sub

Function ToString(vJSON As Variant) As String
    
    Select Case VarType(vJSON)
        Case vbObject, Is >= vbArray
            Set oChunks = CreateObject("Scripting.Dictionary")
            ToStringElement vJSON, ""
            oChunks.Remove 0
            ToString = Join(oChunks.Items(), "")
            Set oChunks = Nothing
        Case vbNull
            ToString = "Null"
        Case vbBoolean
            ToString = IIf(vJSON, "True", "False")
        Case Else
            ToString = CStr(vJSON)
    End Select
    
End Function

Private Sub ToStringElement(vElement As Variant, ByVal sIndent As String)
    
    Dim aKeys() As Variant
    Dim i As Long
    
    With oChunks
        Select Case VarType(vElement)
            Case vbObject
                If vElement.Count = 0 Then
                    .Item(.Count) = "''"
                Else
                    .Item(.Count) = vbCrLf
                    aKeys = vElement.Keys
                    For i = 0 To UBound(aKeys)
                        .Item(.Count) = sIndent & aKeys(i) & ": "
                        ToStringElement vElement(aKeys(i)), sIndent & vbTab
                        If Not (i = UBound(aKeys)) Then .Item(.Count) = vbCrLf
                    Next
                End If
            Case Is >= vbArray
                If UBound(vElement) = -1 Then
                    .Item(.Count) = "''"
                Else
                    .Item(.Count) = vbCrLf
                    For i = 0 To UBound(vElement)
                        .Item(.Count) = sIndent & i & ": "
                        ToStringElement vElement(i), sIndent & vbTab
                        If Not (i = UBound(vElement)) Then .Item(.Count) = vbCrLf
                    Next
                End If
            Case vbNull
                .Item(.Count) = "Null"
            Case vbBoolean
                .Item(.Count) = IIf(vElement, "True", "False")
            Case Else
                .Item(.Count) = CStr(vElement)
        End Select
    End With
    
End Sub

Sub ToArray(vJSON As Variant, aRows() As Variant, aHeader() As Variant)
    
    ' Input:
    ' vJSON - Array or Object which contains rows data
    ' Output:
    ' aRows - 2d array representing JSON data
    ' aHeader - 1d array of property names
    
    Dim sName As Variant
    
    Set oHeader = CreateObject("Scripting.Dictionary")
    Select Case VarType(vJSON)
        Case vbObject
            If vJSON.Count > 0 Then
                ReDim aData(0 To vJSON.Count - 1, 0 To 0)
                oHeader("#") = 0
                i = 0
                For Each sName In vJSON
                    aData(i, 0) = "#" & sName
                    ToArrayElement vJSON(sName), ""
                    i = i + 1
                Next
            Else
                ReDim aData(0 To 0, 0 To 0)
            End If
        Case Is >= vbArray
            If UBound(vJSON) >= 0 Then
                ReDim aData(0 To UBound(vJSON), 0 To 0)
                For i = 0 To UBound(vJSON)
                    ToArrayElement vJSON(i), ""
                Next
            Else
                ReDim aData(0 To 0, 0 To 0)
            End If
        Case Else
            ReDim aData(0 To 0, 0 To 0)
            aData(0, 0) = vJSON
    End Select
    aHeader = oHeader.Keys()
    Set oHeader = Nothing
    aRows = aData
    Erase aData
    
End Sub

Private Sub ToArrayElement(vElement As Variant, sFieldName As String)
    
    Dim sName As Variant
    Dim j As Long
    
    Select Case VarType(vElement)
        Case vbObject ' Collection of objects
            For Each sName In vElement
                ToArrayElement vElement(sName), sFieldName & IIf(sFieldName = "", "", "_") & sName
            Next
        Case Is >= vbArray  ' Collection of arrays
            For j = 0 To UBound(vElement)
                ToArrayElement vElement(j), sFieldName & IIf(sFieldName = "", "", "_") & "#" & j
            Next
        Case Else
            If Not oHeader.Exists(sFieldName) Then
                oHeader(sFieldName) = oHeader.Count
                If UBound(aData, 2) < oHeader.Count - 1 Then ReDim Preserve aData(0 To UBound(aData, 1), 0 To oHeader.Count - 1)
            End If
            j = oHeader(sFieldName)
            aData(i, j) = vElement
    End Select
    
End Sub


