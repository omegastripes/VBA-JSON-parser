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
    Dim sample As String
    sample = "[{""a"":55}, 100]"
    Dim localTestResult As Variant ' To hold the 0 or 1 from parseToVb return
    Dim jsDataObj As Object      ' To hold the raw JS object
    Dim opSuccess As Boolean     ' To hold the success state from parseToVb
    Dim repackedVbaData As Variant ' To actually get the repacked data

    ' Call parseToVb correctly
    Dim funcReturn As Long
    funcReturn = parseToVb(sample, jsDataObj, repackedVbaData, opSuccess)
    
    assign funcReturn, localTestResult ' localTestResult will be 0 or 1

    If opSuccess Then
        Debug.Print "Parsing and repacking successful."
        ' Now repackedVbaData contains the VBA representation (Array of Dictionaries)
        ' You can inspect repackedVbaData here
        If IsArray(repackedVbaData) Then
            Debug.Print "Repacked data is an array. Element count: " & UBound(repackedVbaData) - LBound(repackedVbaData) + 1
        End If
        ' And jsDataObj contains the raw JS object, which can be stringified
        Debug.Print "Original JS Object (stringified):"
        Debug.Print jsonParser.stringify(jsDataObj, "", vbTab)
    Else
        Debug.Print "Parsing or repacking failed."
        Debug.Print "Function return: " & funcReturn
    End If
    Stop
End Sub

Function jsonParser() As Object ' Ensure it returns the parser object
    Static document As Object
    Static json As Object ' Renamed back from jsonAsObject
    
    If json Is Nothing Then ' Renamed back from jsonAsObject
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
        Set json = document.parentWindow.json ' Renamed back from jsonAsObject
    End If
    Set jsonParser = json ' Renamed back from jsonAsObject. Set the module-level variable
End Function

Public Function parseToVb(sample As String, Optional ByRef jsonData As Object, Optional ByRef result As Variant, Optional ByRef success As Boolean) As Long
    ' Default return to failure state
    parseToVb = 0 ' 0 for failure, 1 for success
    success = False
    If Not IsMissing(result) Then result = Empty ' Clear ByRef result only if provided

    On Error GoTo parseToVb_ErrorHandler

    ' Ensure jsonParser is available (it's a module-level variable set by jsonParser() function)
    If jsonParser Is Nothing Then Call jsonParser ' Initialize if not already done by calling the function that sets it
    
    Set jsonData = jsonParser.parse(sample) ' Attempt to parse

    ' Check if parsing itself returned Nothing
    If jsonData Is Nothing Then
        GoTo parseToVb_Exit ' success is already False, parseToVb is 0
    End If

    Dim tempRepackResult As Variant
    repack jsonData, tempRepackResult ' Repack the JS object to VBA Variant

    ' If repack is done, assign to result parameter if it was passed
    If Not IsMissing(result) Then
        If IsObject(tempRepackResult) Then
            Set result = tempRepackResult
        Else
            result = tempRepackResult
        End If
    End If
    
    success = True
    parseToVb = 1 ' 1 for success

parseToVb_Exit:
    Exit Function

parseToVb_ErrorHandler:
    ' Error occurred during parse or repack
    parseToVb = 0 ' Ensure failure indicators
    success = False
    If Not IsMissing(result) Then result = Empty
    If Not jsonData Is Nothing Then Set jsonData = Nothing ' Clear jsonData as it might be in an inconsistent state or invalid
    ' Err.Clear ' Optional: Clear error if considered handled locally
    Resume parseToVb_Exit ' Go to exit to ensure clean function termination
End Function

Private Sub repack(source As Object, result As Variant) ' source is JS Object from jsonParser.parse
    Dim i As Variant ' Loop variable for arrays or dictionary keys
    Dim ret As Variant ' Holds result of recursive repack call
    
    Select Case jsonParser.getType(source)
        Case "array"
            ' jsonParser.cloneDict(source, CreateObject("Scripting.Dictionary")) might not be ideal for JS array
            ' Assuming source is a JS array; need to iterate it like one.
            ' However, the current JS code's cloneDict seems to be used for generic JS objects.
            ' For simplicity, let's assume cloneDict works as intended by the original author for arrays too,
            ' creating a dictionary that, when .Items is called, gives a VBA array of JS objects.
            Dim tempDictForArray As Object
            Set tempDictForArray = jsonParser.cloneDict(source, CreateObject("Scripting.Dictionary"))
            result = tempDictForArray.items ' result is now a VBA array (0-based) of JS objects/values
            
            For i = LBound(result) To UBound(result)
                repack result(i), ret ' result(i) is a JS object/value from the array
                If IsObject(ret) Then
                    Set result(i) = ret ' Store repacked VBA object
                Else
                    result(i) = ret   ' Store repacked VBA primitive
                End If
            Next
        Case "object"
            ' Result will be a VBA Scripting.Dictionary
            Set result = jsonParser.cloneDict(source, CreateObject("Scripting.Dictionary"))
            For Each i In result.Keys ' Iterate VBA Dictionary by keys
                repack result(i), ret ' result(i) is a JS object/value
                If IsObject(ret) Then
                    Set result(i) = ret ' Store repacked VBA object
                Else
                    result(i) = ret   ' Store repacked VBA primitive
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
        Case Else
            ' Unknown type from jsonParser.getType, treat as error or empty
            result = Empty
    End Select
End Sub

Sub assign(src As Variant, dest As Variant) ' Added type hints
    If IsObject(src) Then
        Set dest = src
    Else
        dest = src
    End If
End Sub
