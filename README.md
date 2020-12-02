<p align="center"><img src="https://user-images.githubusercontent.com/3822668/71785573-47c72700-3012-11ea-87c0-87d0e0d006f5.png" /></p>

# VBA JSON Parser
[![release](https://img.shields.io/github/release/omegastripes/VBA-JSON-parser.svg?style=flat&logo=github)](https://github.com/omegastripes/VBA-JSON-parser/releases/latest) [![last-commit](https://img.shields.io/github/last-commit/omegastripes/VBA-JSON-parser.svg?style=flat)](https://github.com/omegastripes/VBA-JSON-parser/commits/master) [![downloads](https://img.shields.io/github/downloads/omegastripes/VBA-JSON-parser/total.svg?style=flat)](https://somsubhra.com/github-release-stats/?username=omegastripes&repository=VBA-JSON-parser) [![code-size](https://img.shields.io/github/languages/code-size/omegastripes/VBA-JSON-parser.svg?style=flat)](https://github.com/omegastripes/VBA-JSON-parser) [![language](https://img.shields.io/github/languages/top/omegastripes/VBA-JSON-parser.svg?style=flat)](https://github.com/omegastripes/VBA-JSON-parser/search?l=vba) [![license](https://img.shields.io/github/license/omegastripes/VBA-JSON-parser.svg?style=flat)](https://github.com/omegastripes/VBA-JSON-parser/blob/master/LICENSE) [![gitter](https://img.shields.io/gitter/room/omegastripes/VBA-JSON-parser.svg?style=flat&logo=gitter)](https://gitter.im/omegastripes) [![tweet](https://img.shields.io/twitter/url/http/shields.io.svg?style=social)](https://twitter.com/intent/tweet?text=Easy%20and%20flexible%20JSON%20processing%20for%20VBA%F0%9F%92%A5&url=https://github.com/omegastripes/VBA-JSON-parser&via=omegastripes&hashtags=vba,json,parse,excel)

[Backus-Naur Form](https://en.wikipedia.org/wiki/Backus%E2%80%93Naur_form) JSON Parser based on RegEx for VBA.
## Purpose and Features
- Parsing JSON string to a structure of nested Dictionaries and Arrays. JSON Objects `{}` are represented by Dictionaries, providing `.Count`, `.Exists()`, `.Item()`, `.Items`, `.Keys` properties and methods. JSON Arrays `[]` are the conventional zero-based VB Arrays, so `UBound() + 1` allows to get the number of elements. Such approach makes easy and straightforward access to structure elements (parsing result is returned via variable passed by ref to sub, so that both an array and a dictionary object can be returned).
- Serializing JSON structure with beautification.
- Building 2D Array based on table-like JSON structure.
- Flattening and unflattening JSON structure.
- Serializing JSON structure into [YAML format](https://yaml.org/) string.
- Parser complies with [JSON Standard](http://json.org/).
- Allows few non-stantard features in JSON string parsing: single quoted and unquoted object keys, single quoted strings, capitalised `True`, `False` and `Null` constants, and trailing commas.
- Invulnerable for malicious JS code injections.
## Compatibility
Supported by MS Windows Office 2003+ (Excel, Word, Access, PowerPoint, Publisher, Visio etc.), CorelDraw, AutoCAD and many others applications with hosted VBA. And even VB6.
## Deployment
Start from example project, Excel workbook is available for downloading in the [latest release](https://github.com/omegastripes/VBA-JSON-parser/releases/latest).

Or

Import **JSON.bas** module into the VBA Project for JSON processing. Need to include a reference to **Microsoft Scripting Runtime**.
<details><summary>How to import?</summary>
<p>

Download and save JSON.bas to a file - open [the page with JSON.bas code](https://github.com/omegastripes/VBA-JSON-parser/blob/master/JSON.bas), right-click on Raw button, choose Save link as... (for Chrome):

![download](https://user-images.githubusercontent.com/3822668/52233449-33dde700-28d0-11e9-97b9-f61fd98c16fd.png)

Import JSON.bas into the VBA Project - open Visual Basic Editor by pressing Alt+F11, right-click on Project Tree, choose Import File, select downloaded JSON.bas:

![import](https://user-images.githubusercontent.com/3822668/52232296-31c65900-28cd-11e9-8164-94ca71c06595.png)

Or you may drag'n'drop downloaded JSON.bas from explorer window (or desktop) directly into the VBA Project Tree.

</p>
</details>
<details><summary>How to add reference?</summary>
<p>

Open Visual Basic Editor by pressing Alt+F11, click Menu - Tools - References, scroll down to **Microsoft Scripting Runtime** and check it, press OK:

![add reference](https://user-images.githubusercontent.com/3822668/71650262-ca579a00-2d25-11ea-9701-4c21dc280ad7.png)

### ![attention](https://user-images.githubusercontent.com/3822668/76687641-cd7cd980-6636-11ea-808d-7fd088be307b.png) MS Word Object Library compatibility note
When referencing both **Microsoft Scripting Runtime** and **Microsoft Word Object Library** make sure that **Microsoft Scripting Runtime** located above **Microsoft Word Object Library** in the the list, if not so then ajust the position by clicking Priority arrows to the right of the list.

![Microsoft Scripting Runtime and Microsoft Word Object Library](https://user-images.githubusercontent.com/3822668/76686982-ed110380-6630-11ea-8d6e-3b4cab94b219.png)

Otherwise you have to change all `Dictionary` references to `Scripting.Dictionary` in your VBA code.

</p>
</details>

## Usage
Here is simple example for MS Excel, put the below code into standard module:

```vba
Option Explicit

Sub Test()
    
    Dim sJSONString As String
    Dim vJSON
    Dim sState As String
    Dim vFlat
    
    ' Retrieve JSON response
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", "http://trirand.com/blog/phpjqgrid/examples/jsonp/getjsonp.php?qwery=longorders&rows=1000", True
        .Send
        Do Until .ReadyState = 4: DoEvents: Loop
        sJSONString = .ResponseText
    End With
    ' Parse JSON response
    JSON.Parse sJSONString, vJSON, sState
    ' Check response validity
    Select Case True
        Case sState <> "Object"
            MsgBox "Invalid JSON response"
        Case Not vJSON.Exists("rows")
            MsgBox "JSON contains no rows"
        Case Else
            ' Convert JSON nested rows array to 2D Array and output to worksheet #1
            Output ThisWorkbook.Sheets(1), vJSON("rows")
            ' Flatten JSON
            JSON.Flatten vJSON, vFlat
            ' Convert to 2D Array and output to worksheet #2
            Output ThisWorkbook.Sheets(2), vFlat
            ' Serialize JSON and save to file
            CreateObject("Scripting.FileSystemObject") _
                .OpenTextFile(ThisWorkbook.Path & "\sample.json", 2, True, -1) _
                .Write JSON.Serialize(vJSON)
            ' Convert JSON to YAML and save to file
            CreateObject("Scripting.FileSystemObject") _
                .OpenTextFile(ThisWorkbook.Path & "\sample.yaml", 2, True, -1) _
                .Write JSON.ToYaml(vJSON)
            MsgBox "Completed"
    End Select
    
End Sub

Sub Output(oTarget As Worksheet, vJSON)
    
    Dim aData()
    Dim aHeader()
    
    ' Convert JSON to 2D Array
    JSON.ToArray vJSON, aData, aHeader
    ' Output to target worksheet range
    With oTarget
        .Activate
        .Cells.Delete
        With .Cells(1, 1)
            .Resize(1, UBound(aHeader) - LBound(aHeader) + 1).Value = aHeader
            .Offset(1, 0).Resize( _
                    UBound(aData, 1) - LBound(aData, 1) + 1, _
                    UBound(aData, 2) - LBound(aData, 2) + 1 _
                ).Value = aData
        End With
        .Columns.AutoFit
    End With
    
End Sub
```

## More Examples
You can find some <a href="https://stackoverflow.com/search?q=user%3A2165759+is%3Aanswer+json.bas">usage examples on SO</a>.

## Beta

Here are some drafts being under development and not fully tested, any bugs detected and suggestions on improvement are welcome in [issues](https://github.com/omegastripes/VBA-JSON-parser/issues).

### Extension <kbd> Beta </kbd>

[jsonExt.bas](https://github.com/omegastripes/VBA-JSON-parser/blob/master/Beta/jsonExt.bas). Some functions available as draft to add flexibility to computations and facilitate processing of JSON structure:

**toArray()** - advanced converting JSON structure to 2d array, enhanced with options explicitly set columns names and order in the header and forbid or permit new columns addition.<br>
**filter()** - fetching elements from array or dictionary by conditions, set like `conds = Array(">=", Array("value", ".dimensions.height"), 15)`.<br>
**sort()** - ordering elements of array or dictionary by value of element by path, set like `".dimensions.height"`.<br>
**slice()** - fetching a part of array or dictionary by beginning and ending indexes.<br>
**selectElement()** - fetching an element from JSON structure by path, set like `".dimensions.height"`.<br>
**joinSubDicts()** - merging properties of subdictionaries from one dictionary to another dictionary.<br>
**joinDicts()** - merging properties from one dictionary to another dictionary.<br>
**nestedArraysToArray()** - converting nested 1d arrays representing table data with header array into array of dictionaries.<br>

### JSON To XML DOM converter <kbd> Beta </kbd>

[JSON2XML.bas](https://github.com/omegastripes/VBA-JSON-parser/blob/master/Beta/JSON2XML.bas). Converting JSON string to XML string and loading it into XML DOM (instead of building a structure of dictionaries and arrays) can significantly increase performance for large data sets. Further XML DOM data processing is not yet covered within current version, and can be implemented via DOM methods and XPath.

### Douglas Crockford json2.js implementation for VBA <kbd> Beta </kbd>

**jsJsonParser** parser is essential for parsing large amounts of JSON data in VBA, it promptly parses strings up to 10 MB and even larger. This implementation built on [douglascrockford/JSON-js](https://github.com/douglascrockford/JSON-js/blob/master/json2.js), native JS code runs on IE JScript engine hosted by htmlfile ActiveX. Parser is wrapped into class module to make it possible to instantiate htmlfile object and create environment for JS execution in Class_Initialize event prior to parsing methods call.

There are two methods available to parse JSON string: `parseToJs(sample, success)` and `parseToVb sample, jsJsonData, result, success`, as follows from the names you can parse to native JS entities of JScriptTypeInfo type, or parse to VBA entities which are a structure of nested Dictionaries and Arrays as described in [Purpose and Features](https://github.com/omegastripes/VBA-JSON-parser#purpose-and-features) section. Access to native JS entities is possible using `jsGetProp()` and `jsGetType()` methods. For JS entities processing you have to have at least common knowledge of JavaScript Objects and Arrays.

Also you can parse to JS entities first, then make some processing and finally convert to VBA entities by calling `parseToVb , jsJsonData, result, success` for further utilization. JS entities can be serialized to JSON string by `stringify(jsJsonData, spacer)` method, if you need to serialize VBA entities, then use `JSON.Serialize()` function from [JSON.bas module](https://github.com/omegastripes/VBA-JSON-parser/blob/master/JSON.bas). If you don't want to mess with JS entities, simply use `parseToVb sample, , result, success` method. Note that convertion to VBA entities will take extra time.

There are few examples in jsJsonParser_v0.1.1.xlsm workbook of the [last release](https://github.com/omegastripes/VBA-JSON-parser/releases/)
