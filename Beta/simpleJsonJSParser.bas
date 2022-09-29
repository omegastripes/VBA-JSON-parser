Attribute VB_Name = "simpleJsonJSParser"

' Douglas Crockford json2.js implementation for VBA
' version 2022-09-29
' https://github.com/douglascrockford/JSON-js/blob/master/json2.js
'
' simpleJsonJSParser derived from jsJsonParser (beta) v0.1.2
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

Sub test()
    Dim sample
    sample = "[{""a"":55}, 100]"
    Dim result
    Dim js
    Dim ok
    assign parseToVb(sample, js, , ok), result
    Debug.Print jsonParser.stringify(js, "", vbTab)
    Stop
End Sub

Function jsonParser()
    Static document As Object
    Static json As Object
    If json Is Nothing Then
        Set document = CreateObject("htmlfile")
        document.Write "<meta http-equiv=""x-ua-compatible"" content=""IE=9"" />'"
        document.parentWindow.execScript Replace( _
            "`object`!=typeof JSON&&(JSON={}),function(){`use strict`;function f(t){return 10>t?`0`+t:t}function this_value(){return this.valueOf()}function quote(t){return rx_escapable.lastIndex=" & _
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
            "})),rx_one.test(text.replace(rx_two,`@`).replace(rx_three,`]`).replace(rx_four,``)))return j=eval(`(`+text+`)`),`function`==typeof reviver?walk({``:j},``):j;throw new SyntaxError(`JSON.parse`)})}();var json=JSON;json.GetType=json.getType=function(t){" & _
            "switch(typeof t){case`string`:case`number`:case`boolean`:case`null`:return typeof t;case`object`:if(!t)return`null`;if(`[object Array]`===Object.prototype.toString.apply(t))return`array`}return`object`};json.CloneDict=json.cloneDict=function(t,e){for" & _
            "(var r in t)e.Add(r,t[r]);return e};json.Parse=json.parse;json.Stringify=json.stringify;", _
            "`", """" _
        )
        Set json = document.parentWindow.json
    End If
    Set jsonParser = json
End Function

Public Function parseToVb(sample, Optional jsonData, Optional result, Optional success)
    result = Empty
    success = False
    On Error Resume Next
    Set jsonData = jsonParser.parse(sample)
    If jsonData Is Nothing Then Exit Function
    Dim vbaJsonObject
    repack jsonData, result
    If Err.Number <> 0 Then Exit Function
    parseToVb = 1
    assign result, parseToVb
    success = True
End Function

Private Sub repack(source, result)
    Select Case jsonParser.getType(source)
        Case "array"
            result = jsonParser.cloneDict(source, CreateObject("Scripting.Dictionary")).items
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
            Set result = jsonParser.cloneDict(source, CreateObject("Scripting.Dictionary"))
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


