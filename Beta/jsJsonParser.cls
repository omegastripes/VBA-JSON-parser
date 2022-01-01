VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "jsJsonParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Douglas Crockford json2.js implementation for VBA
' version 2021-01-01
' https://github.com/douglascrockford/JSON-js/blob/master/json2.js
'
' jsJsonParser (beta) v0.1.2
' Copyright (C) 2021 omegastripes
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

Private Sub Class_Initialize()
    
    Set document = CreateObject("htmlfile")
    document.Write "<meta http-equiv=""x-ua-compatible"" content=""IE=9"" />'"
    With document.parentWindow
        .execScript jsCode()
        Set jsonParser = .JSON
        Set getProp = .getProp
        Set getType = .getType
        Set copyDict = .copyDict
    End With
    
End Sub

Public Property Get jsGetProp() ' As JScriptTypeInfo
    
    Set jsGetProp = getProp
    
End Property

Public Property Get jsGetType() ' As JScriptTypeInfo
    
    Set jsGetType = getType
    
End Property

Public Function parseToJs(sample, Optional success)
    
    On Error Resume Next
    Set parseToJs = jsonParser.parse(sample) ': Dim parse ' keep lower case
    success = Err.Number = 0
    If Not success Then
        Set parseToJs = Nothing
    End If
    
End Function

Public Function parseToVb(Optional sample, Optional jsJsonData, Optional result, Optional success)
    
    If Not IsMissing(sample) Then
        Set jsJsonData = parseToJs(sample)
    End If
    On Error Resume Next
    If jsJsonData Is Nothing Then
        success = False
    Else
        Dim vbaJsonObject
        repack jsJsonData, vbaJsonObject
        success = Err.Number = 0
        If success Then
            If IsObject(vbaJsonObject) Then
                Set parseToVb = vbaJsonObject
                Set result = vbaJsonObject
            Else
                parseToVb = vbaJsonObject
                result = vbaJsonObject
            End If
        Else
            parseToVb = Empty
            result = Empty
            Set jsJsonData = Nothing
        End If
    End If
    
End Function

Public Function stringify(jsJsonData, spacer) ' jsJsonData As JScriptTypeInfo
    
    On Error Resume Next
    stringify = jsonParser.stringify(jsJsonData, "", spacer) ': Dim stringify ' keep lower case
    
End Function

Private Sub repack(source, result)
    
    Select Case getType(source)
        Case "array"
            result = copyDict(source, New Dictionary).items
            Dim i
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
            Set result = copyDict(source, New Dictionary)
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

Private Function jsCode()
    
    ' credits
    ' github repo
    ' https://github.com/douglascrockford/JSON-js/blob/master/json2.js
    ' source json2.js 2017-06-12
    ' https://raw.githubusercontent.com/douglascrockford/JSON-js/master/json2.js
    ' js -minifier
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
