Attribute VB_Name = "simpleJsonJSParser"

' Douglas Crockford json2.js implementation for VBA
' version 2022-09-29
' https://github.com/douglascrockford/JSON-js/blob/master/json2.js
'
' simpleJsonJSParser derived from jsJsonParser (beta) v0.2.0
' Copyright (C) 2026 omegastripes
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

Private document As Object
Private jsonParser As Object
Private getProp As Object
Private getType As Object
Private copyDict As Object
Private chunks As Object
Private headerMap As Object
Private grid() As Variant
Private rowIdx As Long
Private tabChar As String
Private newlineChar As String
Private spaceChar As String

Sub test()
    
    Dim sample
    sample = "[{""a"":55}, 100]"
    Dim result
    Dim js
    Dim ok
    assign parseToVb(sample, js, , ok), result
    Debug.Print stringifyJs(js, vbTab)
    Stop
    
End Sub

Private Sub initJsonParser()
    
    If jsonParser Is Nothing Then
        Set document = CreateObject("htmlfile")
        document.Write "<meta http-equiv=""x-ua-compatible"" content=""IE=9"" />'"
        With document.parentWindow
            .execScript jsCode()
            Set jsonParser = .json
            Set getProp = .getProp
            Set getType = .getType
            Set copyDict = .copyDict
        End With
    End If
    
End Sub

Private Function jsCode()
    
    ' credits
    ' github repo
    ' https://github.com/douglascrockford/JSON-js/blob/master/json2.js
    ' source json2.js 2017-06-12
    ' https://raw.githubusercontent.com/douglascrockford/JSON-js/master/json2.js
    ' js-minifier
    ' http://beautifytools.com/javascript-minifier.php
    
    jsCode = Replace( _
        "function getProp(t,e){return t[e]}function getType(t){switch(typeof t){case`string`:case`number`:case`boolean`:case`null`:return typeof t;case`object`:if(!t)return`null`;if(`[object Array]`===Object.prototype.toString.apply(t))return`array`}return`ob" & _
        "ject`}function copyDict(t,e){for(var r in t)e.Add(r,t[r]);return e}`object`!=typeof JSON&&(JSON={}),function(){`use strict`;function f(t){return 10>t?`0`+t:t}function this_value(){return this.valueOf()}function quote(t){return rx_escapable.lastIndex=" & _
        "0,rx_escapable.test(t)?'`'+t.replace(rx_escapable,function(t){var e=meta[t];return`string`==typeof e?e:`\\u`+(`0000`+t.charCodeAt(0).toString(16)).slice(-4)})+'`':'`'+t+'`'}function str(t,e){var r,n,o,u,f,a=gap,i=e[t];switch(i&&`object`==typeof i&&`f" & _
        "unction`==typeof i.toJSON&&(i=i.toJSON(t)),`function`==typeof rep&&(i=rep.call(e,t,i)),typeof i){case`string`:return quote(i);case`number`:return isFinite(i)?String(i):`null`;case`boolean`:case`null`:return String(i);case`object`:if(!i)return`null`;i" & _
        "f(gap+=indent,f=[],`[object Array]`===Object.prototype.toString.apply(i)){for(u=i.length,r=0;u>r;r+=1)f[r]=str(r,i)||`null`;return o=0===f.length?`[]`:gap?`[\n`+gap+f.join(`,\n`+gap)+`\n`+a+`]`:`[`+f.join(`,`)+`]`,gap=a,o}if(rep&&`object`==typeof rep" & _
        ")for(u=rep.length,r=0;u>r;r+=1)`string`==typeof rep[r]&&(n=rep[r],o=str(n,i),o&&f.push(quote(n)+(gap?`: `:`:`)+o));else for(n in i)Object.prototype.hasOwnProperty.call(i,n)&&(o=str(n,i),o&&f.push(quote(n)+(gap?`: `:`:`)+o));return o=0===f.length?`{}`" & _
        ":gap?`{\n`+gap+f.join(`,\n`+gap)+`\n`+a+`}`:`{`+f.join(`,`)+`}`,gap=a,o}}var rx_one=/^[\],:{}\s]*$/,rx_two=/\\(?:[`\\\/bfnrt]|u[0-9a-fA-F]{4})/g,rx_three=/`[^`\\\n\r]*`|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g,rx_four=/(?:^|:|,)(?:\s*\[)+/" & _
        "g,rx_escapable=/[\\`\u0000-\u001f\u007f-\u009f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,rx_dangerous=/[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\u" & _
        "fff0-\uffff]/g;`function`!=typeof Date.prototype.toJSON&&(Date.prototype.toJSON=function(){return isFinite(this.valueOf())?this.getUTCFullYear()+`-`+f(this.getUTCMonth()+1)+`-`+f(this.getUTCDate())+`T`+f(this.getUTCHours())+`:`+f(this.getUTCMinutes()" & _
        ")+`:`+f(this.getUTCSeconds())+`Z`:null},Boolean.prototype.toJSON=this_value,Number.prototype.toJSON=this_value,String.prototype.toJSON=this_value);var gap,indent,meta,rep;`function`!=typeof JSON.stringify&&(meta={`\b`:`\\b`,`  `:`\\t`,`\n`:`\\n`,`\f`:" & _
        "`\\f`,`\r`:`\\r`,'`':'\\`',`\\`:`\\\\`},JSON.stringify=function(t,e,r){var n;if(gap=``,indent=``,`number`==typeof r)for(n=0;r>n;n+=1)indent+=` `;else`string`==typeof r&&(indent=r);if(rep=e,e&&`function`!=typeof e&&(`object`!=typeof e||`number`!=typeo" & _
        "f e.length))throw new Error(`JSON.stringify`);return str(``,{``:t})}),`function`!=typeof JSON.parse&&(JSON.parse=function(text,reviver){function walk(t,e){var r,n,o=t[e];if(o&&`object`==typeof o)for(r in o)Object.prototype.hasOwnProperty.call(o,r)&&(" & _
        "n=walk(o,r),void 0!==n?o[r]=n:delete o[r]);return reviver.call(t,e,o)}var j;if(text=String(text),rx_dangerous.lastIndex=0,rx_dangerous.test(text)&&(text=text.replace(rx_dangerous,function(t){return`\\u`+(`0000`+t.charCodeAt(0).toString(16)).slice(-4)" & _
        "})),rx_one.test(text.replace(rx_two,`@`).replace(rx_three,`]`).replace(rx_four,``)))return j=eval(`(`+text+`)`),`function`==typeof reviver?walk({``:j},``):j;throw new SyntaxError(`JSON.parse`)})}();var json=JSON;", _
        "`", """" _
    )
    
End Function

Function parseToJs(sample, Optional success)
    
    On Error Resume Next
    initJsonParser
    Set parseToJs = jsonParser.parse(sample) ': Dim parse ' keep lower case
    success = Err.Number = 0
    If Not success Then
        Set parseToJs = Nothing
    End If
    
End Function

Function parseToVb(Optional sample, Optional jsData, Optional result, Optional success)
    
    If Not IsMissing(sample) Then
        Set jsData = parseToJs(sample)
    End If
    result = Empty
    success = False
    On Error Resume Next
    If jsData Is Nothing Then
        success = False
    Else
        Dim obj
        repack jsData, obj
        success = Err.Number = 0
        If success Then
            If IsObject(obj) Then
                Set parseToVb = obj
                Set result = obj
            Else
                parseToVb = obj
                result = obj
            End If
        Else
            parseToVb = Empty
            result = Empty
            Set jsData = Nothing
        End If
    End If
    
End Function

Function stringifyJs(jsData, spacer)
    
    On Error Resume Next
    initJsonParser
    stringifyJs = jsonParser.stringify(jsData, "", spacer) ': Dim stringify ' keep lower case
    
End Function

Function stringifyVb(obj, Optional spacer = vbTab)
    
    If spacer = "" Then
        tabChar = ""
        newlineChar = ""
        spaceChar = ""
    Else
        tabChar = spacer
        newlineChar = vbCrLf
        spaceChar = " "
    End If
    Set chunks = CreateObject("Scripting.Dictionary")
    stringifyVbElement obj, ""
    stringifyVb = Join(chunks.Items(), "")
    Set chunks = Nothing
    
End Function

Private Sub stringifyVbElement(element, indent)
    
    Dim keyList, i
    With chunks
        Select Case VarType(element)
            Case vbObject
                If element.Count = 0 Then
                    .Item(.Count) = "{}"
                Else
                    .Item(.Count) = "{" & newlineChar
                    keyList = element.Keys
                    For i = 0 To UBound(keyList)
                        .Item(.Count) = indent & tabChar & """" & keyList(i) & """:" & spaceChar
                        stringifyVbElement element(keyList(i)), indent & tabChar
                        If Not (i = UBound(keyList)) Then .Item(.Count) = ","
                        .Item(.Count) = newlineChar
                    Next
                    .Item(.Count) = indent & "}"
                End If
            Case Is >= vbArray
                If UBound(element) = -1 Then
                    .Item(.Count) = "[]"
                Else
                    .Item(.Count) = "[" & newlineChar
                    For i = 0 To UBound(element)
                        .Item(.Count) = indent & tabChar
                        stringifyVbElement element(i), indent & tabChar
                        If Not (i = UBound(element)) Then .Item(.Count) = ","
                        .Item(.Count) = newlineChar
                    Next
                    .Item(.Count) = indent & "]"
                End If
            Case vbInteger, vbLong
                .Item(.Count) = element
            Case vbSingle, vbDouble
                .Item(.Count) = Replace(element, ",", ".")
            Case vbNull, vbError
                .Item(.Count) = "null"
            Case vbBoolean
                .Item(.Count) = IIf(element, "true", "false")
            Case Else
                .Item(.Count) = """" & escapeJsonString(element) & """"
        End Select
    End With
    
End Sub

Private Function escapeJsonString(s)
    
    escapeJsonString = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(s, _
        "\", "\\"), _
        """", "\"""), _
        "/", "\/"), _
        Chr(8), "\b"), _
        Chr(12), "\f"), _
        vbLf, "\n"), _
        vbCr, "\r"), _
        vbTab, "\t")
    
End Function

Sub toArray(obj, rows, headers)
    
    Set headerMap = CreateObject("Scripting.Dictionary")
    Select Case VarType(obj)
        Case vbObject
            If obj.Count > 0 Then
                ReDim grid(0 To obj.Count - 1, 0 To 0)
                headerMap("#") = 0
                rowIdx = 0
                Dim k
                For Each k In obj.Keys
                    grid(rowIdx, 0) = k
                    toArrayElement obj(k), ""
                    rowIdx = rowIdx + 1
                Next
            Else
                ReDim grid(0 To 0, 0 To 0)
            End If
        Case Is >= vbArray
            If UBound(obj) >= 0 Then
                ReDim grid(0 To UBound(obj), 0 To 0)
                For rowIdx = 0 To UBound(obj)
                    toArrayElement obj(rowIdx), ""
                Next
            Else
                ReDim grid(0 To 0, 0 To 0)
            End If
        Case Else
            ReDim grid(0 To 0, 0 To 0)
            grid(0, 0) = obj
    End Select
    headers = headerMap.Keys()
    Set headerMap = Nothing
    rows = grid
    Erase grid
    
End Sub

Private Sub toArrayElement(element, field)
    
    Dim fn, colIdx
    Select Case VarType(element)
        Case vbObject
            For Each fn In element.Keys
                toArrayElement element(fn), field & IIf(field = "", "", ".") & fn
            Next
        Case Is >= vbArray
            For colIdx = 0 To UBound(element)
                toArrayElement element(colIdx), field & "[" & colIdx & "]"
            Next
        Case Else
            If Not headerMap.Exists(field) Then
                headerMap(field) = headerMap.Count
                If UBound(grid, 2) < headerMap.Count - 1 Then ReDim Preserve grid(0 To UBound(grid, 1), 0 To headerMap.Count - 1)
            End If
            colIdx = headerMap(field)
            grid(rowIdx, colIdx) = element
    End Select
    
End Sub

Private Sub repack(source, result)
    
    Dim i
    Select Case getType(source)
        Case "array"
            result = copyDict(source, CreateObject("Scripting.Dictionary")).items
            For i = 0 To UBound(result)
                Dim ret
                repack result(i), ret
                If IsObject(ret) Then
                    Set result(i) = ret
                Else
                    result(i) = ret
                End If
            Next
        Case "object"
            Set result = copyDict(source, CreateObject("Scripting.Dictionary"))
            For Each i In result
                repack result(i), ret
                If IsObject(ret) Then
                    Set result(i) = ret
                Else
                    result(i) = ret
                End If
            Next
        Case "string"
            result = CStr(source)
        Case "number"
            result = CDbl(source)
        Case "boolean"
            result = CBool(source)
        Case "null"
            result = Null
    End Select
    
End Sub

Sub assign(src, dest)
    
    If IsObject(src) Then
        Set dest = src
    Else
        dest = src
    End If
    
End Sub

Function jsGetProp(obj, prop)
    
    assign getProp(obj, prop), jsGetProp
    
End Function

Function jsGetType(obj)
    
    jsGetType = getType(obj)
    
End Function


