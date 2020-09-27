Attribute VB_Name = "jsonExt"
' Extension (beta) v0.1.101 for VBA JSON parser, Backus-Naur form JSON parser based on RegEx v1.7.21
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

Private chunks As Dictionary
Private headerList As Dictionary
Private data2dArray() As Variant
Private i As Long
Private skipNew As Boolean
Private ascend As Boolean

Sub toArray(jsonData As Variant, body() As Variant, head() As Variant, Optional skipNewNames As Boolean = False)
    
    ' Input:
    ' jsonData - Array or Object which contains rows data
    ' head - Empty array or array of explicitly set properties names (property name for properties names is "#")
    ' skipNewNames - Behavior of processing new properties names, uses head only if True
    ' Output:
    ' body - 2d array representing JSON data
    ' head - 1d array of property names
    
    skipNew = skipNewNames
    Set headerList = New Dictionary
    Dim field As Variant
    Dim j As Long
    If safeUBound(head) >= 0 Then
        For Each field In head
            If Not headerList.exists(field) Then headerList(field) = headerList.count
        Next
        j = headerList.count - 1
    Else
        j = 0
    End If
    Select Case VarType(jsonData)
        Case vbObject
            If jsonData.count > 0 Then
                If Not skipNew Then
                    headerList("#") = 0
                End If
                ReDim data2dArray(0 To jsonData.count - 1, 0 To j)
                i = 0
                For Each field In jsonData.keys()
                    If skipNew Then
                        If headerList.exists("#") Then
                            j = headerList("#")
                            data2dArray(i, j) = field
                        End If
                    Else
                        data2dArray(i, 0) = field
                    End If
                    toArrayElement jsonData(field), ""
                    i = i + 1
                Next
            Else
                ReDim data2dArray(0 To 0, 0 To j)
            End If
        Case Is >= vbArray
            If UBound(jsonData) >= 0 Then
                ReDim data2dArray(0 To UBound(jsonData), 0 To j)
                For i = 0 To UBound(jsonData)
                    toArrayElement jsonData(i), ""
                Next
            Else
                ReDim data2dArray(0 To 0, 0 To j)
            End If
        Case Else
            ReDim data2dArray(0 To 0, 0 To j)
            data2dArray(0, 0) = jsonData
    End Select
    head = headerList.keys()
    Set headerList = Nothing
    body = data2dArray
    Erase data2dArray
    
End Sub

Private Sub toArrayElement(element As Variant, fieldName As String)
    
    Dim j As Long
    Select Case VarType(element)
        Case vbObject ' Collection of objects
            Dim field As Variant
            For Each field In element.keys()
                toArrayElement element(field), fieldName & IIf(fieldName = "", "", ".") & field
            Next
        Case Is >= vbArray  ' Collection of arrays
            For j = 0 To UBound(element)
                toArrayElement element(j), fieldName & "[" & j & "]"
            Next
        Case Else
            If skipNew Then
                If headerList.exists(fieldName) Then
                    j = headerList(fieldName)
                    data2dArray(i, j) = element
                End If
            Else
                If Not headerList.exists(fieldName) Then
                    headerList(fieldName) = headerList.count
                    If UBound(data2dArray, 2) < headerList.count - 1 Then ReDim Preserve data2dArray(0 To UBound(data2dArray, 1), 0 To headerList.count - 1)
                End If
                j = headerList(fieldName)
                data2dArray(i, j) = element
            End If
    End Select
    
End Sub

Public Function flatten(jsonData)
    
    Set chunks = New Dictionary
    flattenElement jsonData, ""
    Set flatten = chunks
    Set chunks = Nothing
    
End Function

Private Sub flattenElement(element As Variant, property As String)
    
    
    Select Case True
        Case TypeOf element Is Dictionary
            If element.count > 0 Then
                Dim key
                For Each key In element.keys()
                    flattenElement element(key), IIf(property <> "", property & "." & key, key)
                Next
            End If
        Case IsObject(element)
        Case IsArray(element)
            Dim i As Long
            For i = 0 To UBound(element)
                flattenElement element(i), property & "[" & i & "]"
            Next
        Case Else
            chunks(property) = element
    End Select
    
End Sub

Public Sub nestedArraysToArray(body, head, data, success)
    
    ' Input:
        ' body - nested 1d arrays representing table data
        ' head - 1d array of property names
    ' Output:
        ' data - resulting array with JSON data
        ' success - false if head and nested array sizes not match
    
    Dim buffer
    buffer = Array()
    Dim i
    For i = 0 To UBound(body)
        Dim entry
        entry = body(i)
        success = UBound(head) = UBound(entry)
        If Not success Then Exit For
        Dim props
        Set props = New Dictionary
        Dim j
        For j = 0 To UBound(head)
            props(head(j)) = entry(j)
        Next
        pushItem buffer, props
    Next
    data = buffer
    
End Sub

Public Sub filterElements(root, conditions, inclusive, result, success)
    
    ' filtering elements of root array or object
    ' input:
        ' root - source array or object which elements to be filtered
        ' conditions - condition array or nested condition arrays that finally must be evaluated as boolean
        ' inclusive - element will be added to result if
            ' evaluation is true and inclusive is true
            ' evaluation is false or n/a and inclusive is false
    ' output:
        ' result - array or object with filtered elements
        ' success - false if root isn't array or object
    ' condition array description:
    ' supported operations
        ' retrieve element by path relative to root as scalar value or any other JSON entity
            ' <condition> = Array("value", <path>)
            ' <path> - string, expression in JS format, path relative to element of root array or object
            ' example: Array("value", "[0].volume")
        ' return count of elements in root by path as number
            ' <condition> = Array("count", <path>)
            ' <path> - string, expression in JS format, path relative to element of root array or object
            ' example: Array("count", ".shape.points")
        ' check if element exists by path relative to root as boolean
            ' <condition> = Array("exists", <path>)
            ' <path> - string, expression in JS format, path relative to element of root array or object
            ' example: Array("exists", ".items[0].restrictions")
        ' compare two values and return result as boolean
            ' <condition> = Array(<operation>, <expression>, <expression>)
            ' <operation> - string: "=", "<>", ">", ">=", "<", "<="
            ' <expression> - scalar, or nested <condition> evaluated as scalar
            ' example: Array(">=", ".data.volume", 100) - evaluation is true if .data.volume >= 100
        ' check if value belongs to interval specified by two values and return result as boolean
            ' <condition> = Array(<operation>, <expression>, <expression>, <expression>)
            ' <operation> - string: "[]", "[)", "(]", "()"
            ' square brackets mean the end point is included, round parentheses mean it's excluded
            ' <expression> - scalar, or nested <condition> evaluated as scalar
            ' example: Array("[]", ".", 0, 100) - evaluation is true if element value itself >= 0 and <= 100
        ' boolean unary
            ' <condition> = Array("not", <expression>)
            ' <expression> - boolean, or nested <condition> evaluated as boolean
            ' example: Array("not", Array("[]", ".", 0, 100)) - evaluation is true if element value itself < 0 or > 100
            ' "not" operation can be concatenated with any other operation which returns boolean
            ' examples:
            ' Array("not exists", ".items[0].restrictions")
            ' Array("not >=", ".data.volume", 100)
            ' Array("not ()", ".", 0, 100)
        ' boolean binary
            ' <condition> = Array(<operation>, <expression>, <expression>)
            ' <operation> - string: "or", "and", "xor"
            ' <expression> - boolean, or nested <condition> evaluated as boolean
            ' example: Array("and", Array("[]", ".volume", 0, 100), Array(">", ".height", 50))
            ' "or", "and" operations actually accept > 2 arguments, "or" provides lazy evaluation
            ' <condition> = Array(<operation>, <expression>, <expression>, ...)
            ' example:
            ' Array("and", Array("[]", Array("value", ".volume"), 0, 100), Array(">", Array("value", ".height"), 50), Array("<", Array("count", ".specification.items"), 10))
            ' the same example serialized:
            '   [
            '       "and",
            '       [
            '           "[]",
            '           [
            '               "value",
            '               ".volume"
            '           ],
            '           0,
            '           100
            '       ],
            '       [
            '           ">",
            '           [
            '               "value",
            '               ".height"
            '           ],
            '           50
            '       ],
            '       [
            '           "<",
            '           [
            '               "count",
            '               ".specification.items"
            '           ],
            '           10
            '       ]
            '   ]
    
    Dim data
    Dim k
    Dim ret
    Dim decision
    Dim ok
    If IsArray(root) Then
        data = Array()
        For k = 0 To safeUBound(root)
            evaluateExpression root(k), conditions, ret, ok
            decision = False
            Select Case False
                Case ok
                Case VarType(ret) = vbBoolean
                Case ret
                Case Else
                    decision = True
            End Select
            If decision Xor Not inclusive Then
                pushItem data, root(k)
            End If
        Next
        result = data
        success = True
    ElseIf TypeOf root Is Dictionary Then
        Set data = New Dictionary
        For Each k In root.keys()
            evaluateExpression root(k), conditions, ret, ok
            decision = False
            Select Case False
                Case ok
                Case VarType(ret) = vbBoolean
                Case ret
                Case Else
                    decision = True
            End Select
            If decision Xor Not inclusive Then
                If IsObject(root(k)) Then
                    Set data(k) = root(k)
                Else
                    data(k) = root(k)
                End If
            End If
        Next
        Set result = data
        success = True
    Else
        success = False
    End If
    
End Sub

Public Sub groupElements(root, path, full, result, success)
    
    ' grouping elements of root array or object
    ' input:
        ' root - source array or object which elements to be grouped
        ' path - string, expression in JS format, path relative to element of root array or object to entity it grouped by, or array of path components
        ' full - true to create null group for elements having no specified path
    ' output:
        ' result - dictionary with sorted elements with group names as keys
        ' success - false if root isn't array or object
    
    Dim k
    Dim entry
    Dim exists
    If IsArray(root) Then
        Set result = New Dictionary
        Dim buffer
        buffer = Array()
        For k = 0 To safeUBound(root)
            selectElement root(k), path, entry, exists
            If exists Or full Then
                If Not exists Then
                    entry = Null
                End If
                If Not result.exists(entry) Then
                    result(entry) = result.count
                    jsonExt.pushItem buffer, Array()
                End If
                jsonExt.pushItem buffer(result(entry)), root(k)
            End If
        Next
        For Each k In result.keys()
            result(k) = buffer(result(k))
        Next
        success = True
    ElseIf TypeOf root Is Dictionary Then
        Set result = New Dictionary
        For Each k In root.keys()
            selectElement root(k), path, entry, exists
            If exists Or full Then
                If Not exists Then
                    entry = Null
                End If
                If Not result.exists(entry) Then
                    Set result(entry) = New Dictionary
                End If
                If IsObject(root(k)) Then
                    Set result(entry)(k) = root(k)
                Else
                    result(entry)(k) = root(k)
                End If
            End If
        Next
        success = True
    Else
        success = False
    End If
    
End Sub

Private Sub evaluateExpression(root, expr, result, success)
    
    Dim operation
    operation = LCase(expr(0))
    Dim value1
    Dim value2
    Dim value3
    Dim ok
    If Left(operation, 4) = "not " Then
        Dim subexpr
        subexpr = expr
        subexpr(0) = Mid(operation, 5)
        evaluateExpression root, subexpr, value1, ok
        If ok And VarType(value1) = vbBoolean Then
            result = Not value1
            success = True
            Exit Sub
        End If
    End If
    success = False
    Dim exists
    Select Case operation
        Case ""
        Case "value"
            selectElement root, expr(1), value1, exists
            success = exists
            If Not success Then
                Exit Sub
            End If
            assign value1, result
        Case "count"
            selectElement root, expr(1), value1, exists
            success = exists
            If success Then
                If IsArray(value1) Then
                    result = UBound(value1) + 1
                ElseIf TypeOf root Is Dictionary Then
                    result = value1.count
                Else
                    success = False
                End If
            End If
        Case "exists"
            selectElement root, expr(1), value1, exists
            result = exists
            success = True
        Case "=", "<>", ">", ">=", "<", "<=", "[]", "[)", "(]", "()"
            If isScalar(expr(1)) Then
                value1 = expr(1)
            Else
                evaluateExpression root, expr(1), value1, ok
                If Not (ok And isScalar(value1)) Then
                    Exit Sub
                End If
            End If
            If isScalar(expr(2)) Then
                value2 = expr(2)
            Else
                evaluateExpression root, expr(2), value2, ok
                If Not (ok And isScalar(value2)) Then
                    Exit Sub
                End If
            End If
            Select Case operation
                Case "[]", "[)", "(]", "()"
                    If isScalar(expr(3)) Then
                        value3 = expr(3)
                    Else
                        evaluateExpression root, expr(3), value3, ok
                        If Not (ok And isScalar(value3)) Then
                            Exit Sub
                        End If
                    End If
            End Select
            Select Case operation
                Case "="
                    result = CBool(value1 = value2)
                Case "<>"
                    result = CBool(value1 <> value2)
                Case ">"
                    result = CBool(value1 > value2)
                Case ">="
                    result = CBool(value1 >= value2)
                Case "<"
                    result = CBool(value1 < value2)
                Case "<="
                    result = CBool(value1 <= value2)
                Case "[]"
                    result = CBool((value1 >= value2) And (value1 <= value3))
                Case "[)"
                    result = CBool((value1 >= value2) And (value1 < value3))
                Case "(]"
                    result = CBool((value1 > value2) And (value1 <= value3))
                Case "()"
                    result = CBool((value1 > value2) And (value1 < value3))
            End Select
            success = True
        Case "or", "and"
            value2 = True
            Dim i
            For i = 1 To UBound(expr)
                If VarType(expr(i)) = vbBoolean Then
                    value1 = expr(i)
                Else
                    evaluateExpression root, expr(i), value1, ok
                    If Not (ok And VarType(value1) = vbBoolean) Then
                        Exit Sub
                    End If
                End If
                If value1 And operation = "or" Then
                    result = True
                    success = True
                    Exit Sub
                End If
                value2 = value2 And value1
                If Not value2 And operation = "and" Then
                    result = False
                    success = True
                    Exit Sub
                End If
            Next
            If operation = "or" Then
                result = False
            Else
                result = True
            End If
            success = True
        Case "xor", "not"
            If VarType(expr(1)) = vbBoolean Then
                value1 = expr(1)
            Else
                evaluateExpression root, expr(1), value1, ok
                If Not (ok And VarType(value1) = vbBoolean) Then
                    Exit Sub
                End If
            End If
            If operation = "not" Then
                result = Not value1
            Else
                If VarType(expr(2)) = vbBoolean Then
                    value2 = expr(2)
                Else
                    evaluateExpression root, expr(2), value2, ok
                    If Not (ok And VarType(value2) = vbBoolean) Then
                        Exit Sub
                    End If
                End If
                result = value1 And value2
            End If
            success = True
    End Select
    
End Sub

Public Sub sort(root, path, ascending, result)
    
    ' sorting elements of root array or object
    ' input:
        ' root - source array or object which elements to be sorted
        ' path - string, expression in JS format, path relative to element of root array or object to entity it sorted by, or array of path components
        ' ascending - sorting direction
    ' output:
        ' result - array or object with sorted elements
    
    ascend = ascending
    Dim sample()
    sample = Array()
    Dim index()
    index = Array()
    Dim data
    Dim last
    Dim k
    Dim entry
    Dim exists
    Dim i
    If IsArray(root) Then
        data = Array()
        last = safeUBound(root)
        If last >= 0 Then
            ReDim sample(last)
            ReDim index(last)
            ReDim data(last)
            For k = 0 To last
                index(k) = k
                sample(k) = Null
                selectElement root(k), path, entry, exists
                Select Case False
                    Case exists
                    Case Not IsEmpty(entry)
                    Case isScalar(entry)
                    Case Else
                        sample(k) = entry
                End Select
            Next
            quickSortIndex sample, index
            For k = 0 To last
                i = index(k)
                If IsObject(root(i)) Then
                    Set data(k) = root(i)
                Else
                    data(k) = root(i)
                End If
            Next
        End If
        result = data
    ElseIf TypeOf root Is Dictionary Then
        Set data = New Dictionary
        Dim keys
        keys = root.keys()
        last = UBound(keys)
        If last >= 0 Then
            ReDim sample(last)
            ReDim index(last)
            For k = 0 To last
                index(k) = k
                sample(k) = Null
                selectElement root(keys(k)), path, entry, exists
                Select Case False
                    Case exists
                    Case Not IsEmpty(entry)
                    Case isScalar(entry)
                    Case Else
                        sample(k) = entry
                End Select
            Next
            quickSortIndex sample, index
            For k = 0 To last
                i = index(k)
                Dim key
                key = keys(i)
                Dim temp
                If IsObject(root(key)) Then
                    Set temp = root(key)
                    Set data(key) = temp
                Else
                    temp = root(key)
                    data(key) = temp
                End If
            Next
        End If
        Set result = data
    Else
        assign root, result
    End If
    
End Sub

Private Sub quickSortIndex(sample, index)
    
    ' https://rosettacode.org/wiki/Sorting_algorithms/Quicksort
    Dim last As Long
    last = UBound(index)
    If last > 0 Then
        Dim ltArray
        ltArray = Array()
        Dim eqArray
        eqArray = Array()
        Dim gtArray
        gtArray = Array()
        Dim p As Long
        p = Int((last + 1) / 2)
        Dim pivot
        pivot = sample(index(p))
        Dim i As Long
        For i = 0 To last
            Dim elt
            elt = sample(index(i))
            Dim gtCheck
            Dim ltCheck
            If ascend Then
                gtCheck = elt > pivot
                ltCheck = elt < pivot
            Else
                gtCheck = elt < pivot
                ltCheck = elt > pivot
            End If
            If gtCheck Then
                ReDim Preserve gtArray(UBound(gtArray) + 1)
                gtArray(UBound(gtArray)) = index(i)
            ElseIf ltCheck Then
                ReDim Preserve ltArray(UBound(ltArray) + 1)
                ltArray(UBound(ltArray)) = index(i)
            ElseIf elt = pivot Then
                ReDim Preserve eqArray(UBound(eqArray) + 1)
                eqArray(UBound(eqArray)) = index(i)
            Else
                If Not IsNull(pivot) Then ' null > pivot
                    ReDim Preserve gtArray(UBound(gtArray) + 1)
                    gtArray(UBound(gtArray)) = index(i)
                ElseIf Not IsNull(elt) Then ' elt < null
                    ReDim Preserve ltArray(UBound(ltArray) + 1)
                    ltArray(UBound(ltArray)) = index(i)
                Else ' null = null
                    ReDim Preserve eqArray(UBound(eqArray) + 1)
                    eqArray(UBound(eqArray)) = index(i)
                End If
            End If
        Next
        quickSortIndex sample, ltArray
        quickSortIndex sample, gtArray
        p = 0
        For i = 0 To UBound(ltArray)
            index(p) = ltArray(i)
            p = p + 1
        Next
        For i = 0 To UBound(eqArray)
            index(p) = eqArray(i)
            p = p + 1
        Next
        For i = 0 To UBound(gtArray)
            index(p) = gtArray(i)
            p = p + 1
        Next
    End If
    
End Sub

Public Sub selectElement(root, path, entry, exists)
    
    ' retrieve entity from root array or object by relative path
    ' input:
        ' root - source array or object entity to be retrieved from
        ' path - string, expression in JS format, path relative to root array or object, or array of path components
    ' output:
        ' path - array of path components
        ' entry - destination entity retrieved from root by relative path
        ' exists - return false if destination entity doesn't exists or path is invalid
    
    Dim elt
    Dim i
    If Not IsArray(path) Then
        Dim parts
        If path = "" Then
            parts = Array()
            exists = True
        Else
            Dim elts
            elts = Split(Replace(Replace(Replace(path, ".", "|."), "[", "|["), "|.|[", "|.["), "|")
            ReDim parts(UBound(elts) - 1)
            If elts(0) <> "" Then Exit Sub
            For i = 1 To UBound(elts)
                exists = False
                elt = elts(i)
                If Left(elt, 1) = "." Then
                    parts(i - 1) = Mid(elt, 2)
                ElseIf Left(elt, 1) = "[" And Right(elt, 1) = "]" Then
                    elt = Mid(elt, 2, Len(elt) - 2)
                    If IsNumeric(elt) Then
                        parts(i - 1) = CLng(elt)
                    Else
                        Exit For
                    End If
                Else
                    Exit For
                End If
                exists = True
            Next
        End If
        If Not exists Then Exit Sub
        path = parts
    End If
    assign root, entry
    exists = True
    For i = 0 To UBound(path)
        exists = False
        elt = path(i)
        If elt = "" Then
            exists = True
            Exit For
        End If
        If IsArray(entry) Then
            If Not VarType(elt) = vbLong Then Exit For
            If elt < LBound(entry) Or elt > UBound(entry) Then Exit For
        ElseIf TypeOf entry Is Dictionary Then
            If Not entry.exists(elt) Then Exit For
        Else
            Exit For
        End If
        If IsObject(entry(elt)) Then
            Set entry = entry(elt)
        Else
            entry = entry(elt)
        End If
        exists = True
    Next
    
End Sub

Public Sub joinSubDicts(acc, src, Optional addNew = True)
    
    If Not (TypeOf acc Is Dictionary And TypeOf src Is Dictionary) Then
        Exit Sub
    End If
    Dim key
    For Each key In src.keys()
        If TypeOf src(key) Is Dictionary Then
            Dim srcSubDict
            Set srcSubDict = src(key)
            Dim accSubDict
            Set accSubDict = Nothing
            If acc.exists(key) Then
                If TypeOf acc(key) Is Dictionary Then
                    Set accSubDict = acc(key)
                End If
            End If
            If accSubDict Is Nothing Then
               Set accSubDict = New Dictionary
               Set acc(key) = accSubDict
            End If
            joinDicts accSubDict, srcSubDict, addNew
        End If
    Next
    
End Sub

Public Sub joinDicts(acc, src, Optional addNew = True)
    
    If Not (TypeOf acc Is Dictionary And TypeOf src Is Dictionary) Then
        Exit Sub
    End If
    Dim key
    Dim temp
    If addNew Then
        For Each key In src.keys()
            If IsObject(src(key)) Then
                Set temp = src(key)
                Set acc(key) = temp
            Else
                temp = src(key)
                acc(key) = temp
            End If
        Next
    Else
        For Each key In src.keys()
            If acc.exists(key) Then
                If IsObject(src(key)) Then
                    Set temp = src(key)
                    Set acc(key) = temp
                Else
                    temp = src(key)
                    acc(key) = temp
                End If
            End If
        Next
    End If
    
End Sub

Public Sub slice(src, Optional result, Optional ByVal a, Optional ByVal b)
    
    Dim m As Long
    If IsArray(src) Then
        m = UBound(src)
    ElseIf TypeOf src Is Dictionary Then
        m = src.count - 1
    End If
    If IsMissing(a) Then
        a = 0
    End If
    If IsMissing(b) Then
        b = m
    End If
    Dim temp
    Dim i
    Dim d As Long
    If Not (IsNumeric(a) And IsNumeric(b)) Then
        If Not IsMissing(result) Then
            assign src, result
        End If
        Exit Sub
    End If
    Dim void As Boolean
    Dim full As Boolean
    If a < 0 And b < 0 Or a > m And b > m Then
        void = True
    ElseIf a = 0 And b = m Then
        full = True
    Else
        If a < 0 Then
            a = 0
        ElseIf a > m Then
            a = m
        End If
        If b < 0 Then
            b = 0
        ElseIf b > m Then
            b = m
        End If
    End If
    If IsArray(src) Then
        If void Then
            temp = Array()
        ElseIf full Then
            temp = src
        ElseIf a = 0 Then
            temp = src
            ReDim Preserve temp(b)
        Else
            ReDim temp(Abs(b - a))
            Dim j
            j = 0
            d = IIf(a > b, -1, 1)
            For i = a To b Step d
                assign src(i), temp(j)
                j = j + 1
            Next
        End If
        If IsMissing(result) Then
            src = temp
        Else
            result = temp
        End If
    ElseIf TypeOf src Is Dictionary Then
        If void Then
            Set temp = New Dictionary
            temp.CompareMode = src.CompareMode
        ElseIf full Then
            Set temp = jsonExt.cloneDictionary(src)
        Else
            Set temp = New Dictionary
            temp.CompareMode = src.CompareMode
            Dim keys
            keys = src.keys()
            d = IIf(a > b, -1, 1)
            For i = a To b Step d
                If IsObject(src(keys(i))) Then
                    Set temp(keys(i)) = src(keys(i))
                Else
                    temp(keys(i)) = src(keys(i))
                End If
            Next
        End If
        If IsMissing(result) Then
            Set src = temp
        Else
            Set result = temp
        End If
    Else
        assign src, result
    End If
    
End Sub

Public Sub getAvg(root, path, avg, sum, qty)
    
    ' compute sum and average of root array or object values by relative path
    ' input:
        ' root - source array or object of entities to be processed
        ' path - string, expression in JS format, path relative to root array or object, or array of path components
    ' output:
        ' path - array of path components
        ' avg - avg value
        ' sum - sum of values
        ' qty - amount of processed entities
    
    avg = Null
    Dim k
    Dim entry
    Dim exists
    sum = 0
    qty = 0
    If IsArray(root) Then
        For k = 0 To safeUBound(root)
            selectElement root(k), path, entry, exists
            If exists Then
                If IsNumeric(entry) Then
                    qty = qty + 1
                    sum = sum + CDbl(entry)
                End If
            End If
        Next
    ElseIf TypeOf root Is Dictionary Then
        For Each k In root.keys()
            selectElement root(k), path, entry, exists
            If exists Then
                If IsNumeric(entry) Then
                    qty = qty + 1
                    sum = sum + CDbl(entry)
                End If
            End If
        Next
    End If
    If qty > 0 Then
        avg = sum / qty
    End If
    
End Sub

Public Sub getMax(root, path, key, ret, qty)
    
    ' retrieve entity from root array or object having max value by relative path
    ' input:
        ' root - source array or object entity to be retrieved from
        ' path - string, expression in JS format, path relative to root array or object, or array of path components
    ' output:
        ' path - array of path components
        ' key - max value entity key
        ' ret - max value
        ' qty - amount of processed entities
    
    ret = Null
    Dim k
    Dim entry
    Dim exists
    Dim e
    qty = 0
    If IsArray(root) Then
        For k = 0 To safeUBound(root)
            selectElement root(k), path, entry, exists
            If exists Then
                If IsNumeric(entry) Then
                    e = CDbl(entry)
                    qty = qty + 1
                    If ret > e Then
                    Else
                        ret = e
                        key = k
                    End If
                End If
            End If
        Next
    ElseIf TypeOf root Is Dictionary Then
        For Each k In root.keys()
            selectElement root(k), path, entry, exists
            If exists Then
                If IsNumeric(entry) Then
                    e = CDbl(entry)
                    qty = qty + 1
                    If ret > e Then
                    Else
                        ret = e
                        key = k
                    End If
                End If
            End If
        Next
    End If
    
End Sub

Public Sub getMin(root, path, key, ret, qty)
    
    ' retrieve entity from root array or object having min value by relative path
    ' input:
        ' root - source array or object entity to be retrieved from
        ' path - string, expression in JS format, path relative to root array or object, or array of path components
    ' output:
        ' path - array of path components
        ' key - min value entity key
        ' ret - min value
        ' qty - amount of processed entities
    
    Dim k
    Dim entry
    Dim exists
    ret = Null
    Dim e
    qty = 0
    If IsArray(root) Then
        For k = 0 To safeUBound(root)
            selectElement root(k), path, entry, exists
            If exists Then
                If IsNumeric(entry) Then
                    e = CDbl(entry)
                    qty = qty + 1
                    If ret < e Then
                    Else
                        ret = e
                        key = k
                    End If
                End If
            End If
        Next
    ElseIf TypeOf root Is Dictionary Then
        For Each k In root.keys()
            selectElement root(k), path, entry, exists
            If exists Then
                If IsNumeric(entry) Then
                    e = CDbl(entry)
                    qty = qty + 1
                    If ret < e Then
                    Else
                        ret = e
                        key = k
                    End If
                End If
            End If
        Next
    End If
    
End Sub

Function safeUBound(a)
    
    safeUBound = -1
    On Error Resume Next
    safeUBound = UBound(a)
    Err.Clear
    
End Function

Function isScalar(v) As Boolean
    
    Select Case VarType(v)
        Case vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbEmpty, vbInteger, vbLong, vbSingle, vbString
            isScalar = True
        Case Else
            isScalar = False
    End Select
    
End Function

Function cloneDictionary(srcDict)
    
    Dim destDict As Dictionary
    Set destDict = New Dictionary
    If TypeOf srcDict Is Dictionary Then
        If Not srcDict Is Nothing Then
            destDict.CompareMode = srcDict.CompareMode
            Dim key
            Dim temp
            For Each key In srcDict.keys()
                If IsObject(srcDict(key)) Then
                    
                    Set temp = srcDict(key)
                    Set destDict(key) = temp
                Else
                    temp = srcDict(key)
                    destDict(key) = temp
                End If
            Next
        End If
    End If
    Set cloneDictionary = destDict
    
End Function

Sub pushItem( _
        destArray, _
        sourceElement, _
        Optional optionAppend As Boolean = True, _
        Optional optionNestArrays As Boolean = True _
    )
    
    ' not optionAppend => create array
    ' sourceElement array and optionAppend => do not create array
    ' sourceElement not array and optionAppend => create array with single elt
    Select Case True
        Case Not optionAppend Or IsEmpty(destArray)
            destArray = Array()
        Case Not IsArray(destArray)
            destArray = Array(destArray)
    End Select
    If IsArray(sourceElement) And Not optionNestArrays Then
        Dim n As Long
        Dim j As Long
        Dim i As Long
        n = UBound(destArray)
        ReDim Preserve destArray(LBound(destArray) To n + UBound(sourceElement) - LBound(sourceElement) + 1)
        j = 1
        For i = LBound(sourceElement) To UBound(sourceElement)
            assign sourceElement(i), destArray(n + j)
            j = j + 1
        Next
    Else
        ReDim Preserve destArray(LBound(destArray) To UBound(destArray) + 1)
        assign sourceElement, destArray(UBound(destArray))
    End If
    
End Sub

Sub assign(source, dest)
    
    If IsObject(source) Then
        Set dest = source
    Else
        dest = source
    End If
    
End Sub
