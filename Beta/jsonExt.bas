Attribute VB_Name = "jsonExt"
' Extension (beta) v0.1 for VBA JSON parser, Backus-Naur form JSON parser based on RegEx v1.7.03
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
    
    Dim field As Variant
    Dim j As Long
    
    skipNew = skipNewNames
    Set headerList = New Dictionary
    If safeUBound(head) >= 0 Then
        For Each field In head
            If Not headerList.exists(field) Then headerList(field) = headerList.Count
        Next
        j = headerList.Count - 1
    Else
        j = 0
    End If
    Select Case VarType(jsonData)
        Case vbObject
            If jsonData.Count > 0 Then
                If Not skipNew Then
                    headerList("#") = 0
                End If
                ReDim data2dArray(0 To jsonData.Count - 1, 0 To j)
                i = 0
                For Each field In jsonData.keys
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
    
    Dim field As Variant
    Dim j As Long
    
    Select Case VarType(element)
        Case vbObject ' Collection of objects
            For Each field In element.keys
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
                    headerList(fieldName) = headerList.Count
                    If UBound(data2dArray, 2) < headerList.Count - 1 Then ReDim Preserve data2dArray(0 To UBound(data2dArray, 1), 0 To headerList.Count - 1)
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
    
    Dim key
    Dim i As Long
    
    Select Case True
        Case TypeOf element Is Dictionary
            If element.Count > 0 Then
                For Each key In element.keys
                    flattenElement element(key), IIf(property <> "", property & "." & key, key)
                Next
            End If
        Case IsObject(element)
        Case isArray(element)
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
    ' success - completed ok
    
    Dim buffer
    Dim i
    Dim entry
    Dim j
    Dim props
    Dim temp
    Dim num
    Dim itv
    
    buffer = Array()
    For i = 0 To UBound(body)
        entry = body(i)
        success = UBound(head) = UBound(entry)
        If Not success Then Exit For
        Set props = New Dictionary
        For j = 0 To UBound(head)
            props(head(j)) = entry(j)
        Next
        pushItem buffer, props
    Next
    data = buffer
    
End Sub

Public Sub filter(root, conds, inclusive, result, success)
    
    ' filtering
    ' condition operations
    ' returns <boolean>
    '   exists, <path>
    '   = | <> | > | >= | < | <=, <value>, <value>
    '   between | between[) | between(] | between(), <value>, <value>, <value>
    '   in, <value>, <value> .. <value>
    '   or | and | xor | not, <value>, <value> .. <value>
    ' returns <value>
    '   count, <path>
    ' returns <value> | <entity>
    '   value, <path>
    
    Dim data
    Dim k
    Dim decision
    Dim ok
    
    If isArray(root) Then
        data = Array()
        For k = 0 To safeUBound(root)
            evaluateCondition root(k), conds, decision, ok
            If ok And VarType(decision) = vbBoolean Then
                If decision Xor Not inclusive Then
                    pushItem data, root(k)
                End If
            End If
        Next
        result = data
        success = True
    ElseIf TypeOf root Is Dictionary Then
        Set data = New Dictionary
        For Each k In root.keys
            evaluateCondition root(k), conds, decision, ok
            If ok And VarType(decision) = vbBoolean Then
                If decision Xor Not inclusive Then
                    If IsObject(root(k)) Then
                        Set data(k) = root(k)
                    Else
                        data(k) = root(k)
                    End If
                End If
            End If
        Next
        Set result = data
        success = True
    Else
        success = False
    End If
    
End Sub

Private Sub evaluateCondition(root, conds, result, success)
    
    Dim operation
    Dim subconds
    Dim value1
    Dim value2
    Dim value3
    Dim ok
    Dim exists
    Dim i
    
    operation = LCase(conds(0))
    If Left(operation, 4) = "not " Then
        subconds = conds
        subconds(0) = Mid(operation, 5)
        evaluateCondition root, subconds, value1, ok
        If ok And VarType(value1) = vbBoolean Then
            result = Not value1
            success = True
            Exit Sub
        End If
    End If
    success = False
    Select Case operation
        Case ""
        Case "value"
            selectElement root, conds(1), value1, exists
            success = exists
            If Not success Then
                Exit Sub
            End If
            assign value1, result
        Case "count"
            selectElement root, conds(1), value1, exists
            success = exists
            If success Then
                If isArray(value1) Then
                    result = UBound(value1) + 1
                ElseIf TypeOf root Is Dictionary Then
                    result = value1.Count
                Else
                    success = False
                End If
            End If
        Case "exists"
            selectElement root, conds(1), value1, exists
            result = exists
            success = True
        Case "=", "<>", ">", ">=", "<", "<=", "between", "between[)", "between(]", "between()"
            If isScalar(conds(1)) Then
                value1 = conds(1)
            Else
                evaluateCondition root, conds(1), value1, ok
                If Not (ok And isScalar(value1)) Then
                    Exit Sub
                End If
            End If
            If isScalar(conds(2)) Then
                value2 = conds(2)
            Else
                evaluateCondition root, conds(2), value2, ok
                If Not (ok And isScalar(value2)) Then
                    Exit Sub
                End If
            End If
            Select Case operation
                Case "between", "between[)", "between(]", "between()"
                    If isScalar(conds(3)) Then
                        value3 = conds(3)
                    Else
                        evaluateCondition root, conds(3), value3, ok
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
                Case "between"
                    result = CBool((value1 >= value2) And (value1 <= value3))
                Case "between[)"
                    result = CBool((value1 >= value2) And (value1 < value3))
                Case "between(]"
                    result = CBool((value1 > value2) And (value1 <= value3))
                Case "between()"
                    result = CBool((value1 > value2) And (value1 < value3))
            End Select
            success = True
        Case "or", "and"
            value2 = True
            For i = 1 To UBound(conds)
                If VarType(conds(i)) = vbBoolean Then
                    value1 = conds(i)
                Else
                    evaluateCondition root, conds(i), value1, ok
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
            If VarType(conds(1)) = vbBoolean Then
                value1 = conds(1)
            Else
                evaluateCondition root, conds(1), value1, ok
                If Not (ok And VarType(value1) = vbBoolean) Then
                    Exit Sub
                End If
            End If
            If operation = "not" Then
                result = Not value1
            Else
                If VarType(conds(2)) = vbBoolean Then
                    value2 = conds(2)
                Else
                    evaluateCondition root, conds(2), value2, ok
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
    
    Dim sample()
    Dim index()
    Dim data
    Dim last
    Dim k
    Dim keys
    Dim entry
    Dim exists
    Dim i
    
    ascend = ascending
    sample = Array()
    index = Array()
    If isArray(root) Then
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
        keys = root.keys
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
                If IsObject(root(keys(i))) Then
                    Set data(keys(i)) = root(keys(i))
                Else
                    data(keys(i)) = root(keys(i))
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
    Dim ltArray
    Dim eqArray
    Dim gtArray
    Dim pivot
    Dim elt
    Dim i As Long
    Dim gtCheck
    Dim ltCheck
    Dim p As Long
    
    last = UBound(index)
    If last > 0 Then
        ltArray = Array()
        eqArray = Array()
        gtArray = Array()
        p = Int((last + 1) / 2)
        pivot = sample(index(p))
        For i = 0 To last
            elt = sample(index(i))
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
    
    Dim elts
    Dim parts
    Dim elt
    Dim i
    
    If Not isArray(path) Then
        If path = "" Then
            parts = Array()
            exists = True
        Else
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
        If isArray(entry) Then
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

    Dim key
    Dim accSubDict
    Dim subKey
    Dim srcSubDict
    
    If Not (TypeOf acc Is Dictionary And TypeOf src Is Dictionary) Then
        Exit Sub
    End If
    For Each key In src.keys
        If TypeOf src(key) Is Dictionary Then
            Set srcSubDict = src(key)
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
    If addNew Then
        For Each key In src.keys
            If IsObject(src(key)) Then
                Set acc(key) = src(key)
            Else
                acc(key) = src(key)
            End If
        Next
    Else
        For Each key In src.keys
            If acc.exists(key) Then
                If IsObject(src(key)) Then
                    Set acc(key) = src(key)
                Else
                    acc(key) = src(key)
                End If
            End If
        Next
    End If
    
End Sub

Public Sub slice(src, Optional result, Optional ByVal a, Optional ByVal B)
    
    Dim m As Long
    Dim void As Boolean
    Dim full As Boolean
    Dim temp
    Dim i
    Dim j
    Dim d As Long
    Dim keys
    
    If isArray(src) Then
        m = UBound(src)
    ElseIf TypeOf src Is Dictionary Then
        m = src.Count - 1
    End If
    If IsMissing(a) Then
        a = 0
    End If
    If IsMissing(B) Then
        B = m
    End If
    If Not (IsNumeric(a) And IsNumeric(B)) Then
        If Not IsMissing(result) Then
            assign src, result
        End If
        Exit Sub
    End If
    If a < 0 And B < 0 Or a > m And B > m Then
        void = True
    ElseIf a = 0 And B = m Then
        full = True
    Else
        If a < 0 Then
            a = 0
        ElseIf a > m Then
            a = m
        End If
        If B < 0 Then
            B = 0
        ElseIf B > m Then
            B = m
        End If
    End If
    If isArray(src) Then
        If void Then
            temp = Array()
        ElseIf full Then
            temp = src
        ElseIf a = 0 Then
            temp = src
            ReDim Preserve temp(B)
        Else
            ReDim temp(Abs(B - a))
            j = 0
            d = IIf(a > B, -1, 1)
            For i = a To B Step d
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
            Set temp = cloneDictionary(src)
        Else
            Set temp = New Dictionary
            temp.CompareMode = src.CompareMode
            keys = src.keys
            d = IIf(a > B, -1, 1)
            For i = a To B Step d
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

Public Sub getAvg(root, path, avg, n)
    
    Dim k
    Dim entry
    Dim exists
    Dim s
    
    n = 0
    If isArray(root) Then
        For k = 0 To safeUBound(root)
            selectElement root(k), path, entry, exists
            If exists Then
                If IsNumeric(entry) Then
                    n = n + 1
                    s = s + CDbl(entry)
                End If
            End If
        Next
    ElseIf TypeOf root Is Dictionary Then
        For Each k In root.keys
            selectElement root(k), path, entry, exists
            If exists Then
                If IsNumeric(entry) Then
                    n = n + 1
                    s = s + CDbl(entry)
                End If
            End If
        Next
    End If
    If n > 0 Then
        avg = s / n
    End If
    
End Sub

Function safeUBound(a)
    
    safeUBound = -1
    On Error Resume Next
    safeUBound = UBound(a)
    Err.Clear
    
End Function

Function isScalar(V) As Boolean
    
    Select Case VarType(V)
        Case vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbEmpty, vbInteger, vbLong, vbSingle, vbString
            isScalar = True
        Case Else
            isScalar = False
    End Select
    
End Function

Function cloneDictionary(srcDict)
    
    Dim destDict As Dictionary
    Dim propName
    Dim propValue
    
    Set destDict = New Dictionary
    If TypeOf srcDict Is Dictionary Then
        If Not srcDict Is Nothing Then
            destDict.CompareMode = srcDict.CompareMode
            For Each propName In srcDict.keys
                If IsObject(srcDict(propName)) Then
                    Set propValue = srcDict(propName)
                    Set destDict(propName) = propValue
                Else
                    propValue = srcDict(propName)
                    destDict(propName) = propValue
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
        Case Not isArray(destArray)
            destArray = Array(destArray)
    End Select
    If isArray(sourceElement) And Not optionNestArrays Then
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
