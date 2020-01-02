# VBA JSON Parser
[Backus-Naur Form](https://en.wikipedia.org/wiki/Backus%E2%80%93Naur_form) JSON Parser based on RegEx for VBA.
## Purpose and Features
- Converts JSON string to a structure of nested Dictionaries and Arrays. JSON Objects `{}` are represented by Dictionaries, providing `.Count`, `.Exists()`, `.Item()`, `.Items`, `.Keys` properties and methods. JSON Arrays `[]` are the conventional zero-based VB Arrays, so `UBound()` allows to get the number of elements. Such approach makes easy and straightforward access to structure elements.
- Serialize and beautify JSON structure.
- Builds 2D Array based on table-like JSON structure.
- Flattens JSON structure.
- Serialize JSON structure into [YAML format](https://yaml.org/) string.
- Parser complies with [JSON Standard](http://json.org/).
- Allows few non-stantard features in JSON string parsing: single quoted and unquoted object keys, single quoted strings, capitalised `True`, `False` and `Null` constants, and trailing commas.
- Invulnerable for malicious JS code injections.
## Compatibility
Supported by MS Windows Office 2003+ (Excel, Word, Access, PowerPoint, Publisher, Visio etc.), CorelDraw, AutoCAD and many others applications with hosted VBA. And even VB6.
## Deployment
Start from example project, Excel workbook is available for downloading in the [latest release](https://github.com/omegastripes/VBA-JSON-parser/releases/latest).

Or

Import **JSON.bas** module into the VBA Project for JSON processing. Need to include a reference to "Microsoft Scripting Runtime".
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

Open Visual Basic Editor by pressing Alt+F11, click Menu - Tools - References, scroll down to "Microsoft Scripting Runtime" and check it, OK:

![add reference](https://user-images.githubusercontent.com/3822668/71650262-ca579a00-2d25-11ea-9701-4c21dc280ad7.png)

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

**Extension** (VBA-JSON-parser/Beta/jsonExt.bas)

Some functions available as draft to add flexibility to computations and facilitate processing of JSON structure:

**toArray()** - converts JSON structure to 2d array, enhanced with option explicitly set columns names and order in the header and forbid or permit new columns addition.<br>
**nestedArraysToArray()** - converts nested 1d arrays representing table data with header array into array of dictionaries.<br>
**filter()** - fetches elements from array or dictionary by conditions, set like `conds = Array(">=", Array("value", ".dimensions.height"), 15)`.<br>
**sort()** - orders elements of array or dictionary by value of element by path, set like `".dimensions.height"`.<br>
**selectElement()** - fetch an element from JSON structure by path, set like `".dimensions.height"`.<br>
**joinSubDicts()** - merges properties of subdictionaries from one dictionary to another dictionary.<br>
**joinDicts()** - merges properties from one dictionary to another dictionary.<br>
**slice()** - fetches a part of array or dictionary by beginning and ending indexes.<br>

**JSON To XML DOM converter** (VBA-JSON-parser/Beta/JSON2XML.bas)

Converting JSON string to XML string and loading it into XML DOM (instead of building a structure of dictionaries and arrays) can significantly increase performance for large data sets. Further XML DOM data processing can be implemented via DOM methods and XPath.
