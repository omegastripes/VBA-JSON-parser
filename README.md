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
Import **JSON.bas** module into the VBA Project for JSON processing. No references are necessary in the VBA Project due to late bind.
<details><summary>How to import?</summary>
<p>

Download and save JSON.bas to a file - open the page with JSON.bas code, right-click on Raw button, choose Save link as... (for Chrome):

![download](https://user-images.githubusercontent.com/3822668/52233449-33dde700-28d0-11e9-97b9-f61fd98c16fd.png)

Import JSON.bas into the VBA Project - open Visual Basic Editor by pressing Alt+F11, right-click on Project Tree, choose Import File, select saved JSON.bas:

![import](https://user-images.githubusercontent.com/3822668/52232296-31c65900-28cd-11e9-8164-94ca71c06595.png)

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
