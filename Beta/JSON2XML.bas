Attribute VB_Name = "JSON2XML"
' JSON2XML (beta) v0.1
' Copyright (C) 2015-2020 omegastripes
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

Sub convertJsonToXmlDomTest()
    
    ' convert JSON to XML DOM
    
    ' add references:
    ' Microsoft XML, v6.0
    ' Microsoft Scripting Runtime
    
    Dim content As String
    ' retrieve json
    With New MSXML2.XMLHTTP
        .Open "GET", "http://trirand.com/blog/phpjqgrid/examples/jsonp/getjsonp.php?qwery=longorders&rows=20000", True
        .Send
        Do Until .ReadyState = 4: DoEvents: Loop
        content = .ResponseText
    End With
    saveTextToFile content, ThisWorkbook.Path & "\data.json", "utf-8"
'    ' load json
'    content = loadTextFromFile(ThisWorkbook.Path & "\data.json", "utf-8")
    
    Dim t
    t = Timer
    ' extract strings from json body
    With CreateObject("VBScript.RegExp")
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "(""|')((?:\\\1|(?!\1).)*)\1"
        content = .Replace(content, ChrW(0) & "$2" & ChrW(0)) ' ChrW(0) = vbNullChar
        .pattern = "\b([A-Za-z_]\w*)(?=\s*\:)"
        content = .Replace(content, ChrW(0) & "$1" & ChrW(0))
    End With
    Dim chunks
    chunks = Split(content, ChrW(0))
    Dim strings
    strings = Array()
    If UBound(chunks) > 0 Then
        ReDim strings((UBound(chunks) - 1) \ 2) ' 1 - 0, 3 - 1, 5 - 2
        Dim i As Long ' Explicitly declare i
        For i = 1 To UBound(chunks) Step 2
            strings((i - 1) \ 2) = chunks(i)
            chunks(i) = ChrW(0)
        Next
    End If
    ' unescape json chars and encoding html entities
    content = Join(strings, ChrW(0))
    content = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace( _
        content, _
        "<", "&lt;"), _
        ">", "&gt;"), _
        "&", "&amp;"), _
        "'", "&apos;"), _
        "\""", "&quot;"), _
        "\\", "\" & ChrW(-1)), _
        "\/", "/"), _
        "\b", Chr(8)), _
        "\f", Chr(12)), _
        "\n", vbLf), _
        "\r", vbCr), _
        "\t", vbTab)
    strings = Split(content, "\u")
    ' replace unicode chars
    Dim i As Long ' Declare i again for this loop if not already in scope or reuse
    For i = 1 To UBound(strings)
        Dim u
        u = ChrW(("&H" & Left(strings(i), 4)) * 1)
        strings(i) = u & Mid(strings(i), 5)
    Next
    content = Join(strings, "")
    content = Replace(content, "\" & ChrW(-1), "\")
    strings = Split(content, ChrW(0))
    ' simplify json body
    content = Join(chunks, "")
    With CreateObject("VBScript.RegExp")
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "\s+"
        content = .Replace(content, "")
        .pattern = ",,+"
        content = .Replace(content, ",")
    End With
    ' convert json to xml outline
    content = Replace(content, "[,", "[")
    content = Replace(content, "{,", "{")
    content = Replace(content, ",]", "]")
    content = Replace(content, ",}", "}")
    content = Replace(content, ":" & ChrW(0), """ type=""string"">" & ChrW(0))
    content = Replace(content, ":", """>")
    content = Replace(content, "{" & ChrW(0) & """", "<object><property name=""" & ChrW(0) & """")
    content = Replace(content, "," & ChrW(0) & """", "</property><property name=""" & ChrW(0) & """")
    content = Replace(content, "}", "</property></object>")
    content = Replace(content, "[", "<array><element>")
    content = Replace(content, ",", "</element><element>")
    content = Replace(content, "]", "</element></array>")
    ' insert strings back to xml structure
    chunks = Split(content, ChrW(0))
    For i = 1 To UBound(chunks) ' Reuse i
        chunks(i) = strings(i - 1) & chunks(i)
    Next
    content = Join(chunks, "")
    ' load xml dom
    Dim xml As MSXML2.DOMDocument60
    Set xml = New MSXML2.DOMDocument60
    xml.LoadXML content
    Debug.Print "Elapsed " & Round(Timer - t, 3) & " sec"
    isParseXMLSuccess xml
    '
    ' processing xml dom
    '
    saveTextToFile content, ThisWorkbook.Path & "\result_raw.xml", "utf-8"
    ' beautify xml
    Dim xml2 As MSXML2.DOMDocument60
    Set xml2 = beautifyXML(xml)
    saveTextToFile xml2.xml, ThisWorkbook.Path & "\result_beautified.xml", "utf-8"
    
End Sub

Function beautifyXML(xml As MSXML2.DOMDocument60) As MSXML2.DOMDocument60
    
    Dim writer As New MSXML2.MXXMLWriter60
    Dim reader As New MSXML2.SAXXMLReader60
    Dim strContent As String ' Renamed from content to avoid conflict
    
    writer.Indent = True
    writer.omitXMLDeclaration = True
    With reader
        Set .contentHandler = writer
        Set .dtdHandler = writer
        Set .errorHandler = writer
        .putProperty "http://xml.org/sax/properties/lexical-handler", writer
        .putProperty "http://xml.org/sax/properties/declaration-handler", writer
        .Parse xml
    End With
    strContent = writer.output ' Use strContent
    strContent = IIf(Left(strContent, 6) <> "<?xml ", "<?xml version=""1.0"" encoding=""utf-8"" ?>" & vbCrLf, "") & strContent
    
    ' Load the beautified string into a new DOM to return a DOM object
    Dim outDom As MSXML2.DOMDocument60 ' Declare outDom
    Dim loadSuccess As Boolean         ' Declare loadSuccess
    loadXmlFromString strContent, outDom, loadSuccess ' Pass outDom and loadSuccess
    
    If Not loadSuccess Then
        MsgBox "Error in beautifyXML loading string: " & outDom.parseError.reason
        Set beautifyXML = xml ' Return original DOM on error
    Else
        Set beautifyXML = outDom
    End If
    
End Function

Sub loadXmlFromString(ByVal xmlString As String, ByRef xmlDocument As MSXML2.DOMDocument60, ByRef success As Boolean)
    
    Set xmlDocument = New MSXML2.DOMDocument60
    With xmlDocument
        .validateOnParse = False
        .resolveExternals = False
        .async = False ' Ensure synchronous loading
        .setProperty "ProhibitDTD", False
        .setProperty "SelectionLanguage", "XPath"
        .LoadXML xmlString
        success = (.parseError.ErrorCode = 0)
        If Not success Then
             Debug.Print "loadXmlFromString Error: " & .parseError.reason & " on XML: " & xmlString
        End If
    End With
    
End Sub

Function isParseXMLSuccess(xml As MSXML2.DOMDocument60) As Boolean
    
    With xml.parseError
        isParseXMLSuccess = .ErrorCode = 0
        If Not isParseXMLSuccess Then
            MsgBox _
                "XML parsing error: " & _
                .ErrorCode & ", " & _
                .reason & ", " & _
                "line: " & .Line & ", " & _
                "pos:" & .linepos & ", " & _
                "source: " & .srcText, _
                vbExclamation
        End If
    End With
    
End Function

Function loadTextFromFile(filePath, charset)
    
    With CreateObject("ADODB.Stream")
        .Type = 1 ' TypeBinary
        .Open
        .LoadFromFile filePath
        .Position = 0
        .Type = 2 ' adTypeText
        .charset = charset
        loadTextFromFile = .ReadText
        .Close
    End With
    
End Function

Sub saveTextToFile(content, filePath, charset)
    
    smartCreateFolder CreateObject("Scripting.FileSystemObject").GetParentFolderName(filePath)
    With CreateObject("ADODB.Stream")
        .Type = 2 ' adTypeText
        .Open
        .charset = charset
        .WriteText content
        .Position = 0
        .Type = 1 ' TypeBinary
        .SaveToFile filePath, 2
        .Close
    End With
    
End Sub

Sub smartCreateFolder(folder)
    
    With CreateObject("Scripting.FileSystemObject")
        If Not .FolderExists(folder) Then
            smartCreateFolder .GetParentFolderName(folder)
            .CreateFolder folder
        End If
    End With
    
End Sub
