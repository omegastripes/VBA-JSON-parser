Attribute VB_Name = "jsonExt"
' Extension (beta) v0.1.103 for VBA JSON parser, Backus-Naur form JSON parser based on RegEx v1.7.21
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
Private i As Long ' Module level i, used in toArray and toArrayElement
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
                i = 0 ' Uses module-level i
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
                For i = 0 To UBound(jsonData) ' Uses module-level i
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
                    data2dArray(i, j) = element ' Uses module-level i
                End If
            Else
                If Not headerList.exists(fieldName) Then
                    headerList(fieldName) = headerList.count
                    If UBound(data2dArray, 2) < headerList.count - 1 Then ReDim Preserve data2dArray(0 To UBound(data2dArray, 1), 0 To headerList.count - 1)
                End If
                j = headerList(fieldName)
                data2dArray(i, j) = element ' Uses module-level i
            End If
    End Select
    
End Sub

Public Function flatten(jsonData As Variant) As Dictionary ' Return type specified
    
    Set chunks = New Dictionary
    flattenElement jsonData, ""
    Set flatten = chunks
    Set chunks = Nothing
    
End Function

Private Sub flattenElement(element As Variant, currentProperty As String) ' Renamed parameter
    
    Select Case True
        Case TypeOf element Is Dictionary
            If element.count > 0 Then
                Dim key As Variant
                For Each key In element.keys()
                    flattenElement element(key), IIf(currentProperty <> "", currentProperty & "." & key, key)
                Next
            End If
        Case IsObject(element) ' Could be other types of objects, handle if necessary or ignore
             ' For now, this case does nothing further if not a Dictionary
        Case IsArray(element)
            Dim idx As Long ' Renamed loop variable
            For idx = LBound(element) To UBound(element) ' Use LBound for robustness
                flattenElement element(idx), currentProperty & "[" & idx & "]"
            Next
        Case Else ' Scalar values
            chunks(currentProperty) = element
    End Select
    
End Sub

Public Sub nestedArraysToArray(body As Variant, head As Variant, ByRef data As Variant, ByRef success As Boolean)
    
    ' Input:
        ' body - nested 1d arrays representing table data
        ' head - 1d array of property names
    ' Output:
        ' data - resulting array with JSON data
        ' success - false if head and nested array sizes not match
    
    Dim buffer As Variant
    buffer = Array()
    Dim local_i As Long ' Use local loop variable
    success = False ' Initialize success

    If Not IsArray(body) Or Not IsArray(head) Then Exit Sub ' Basic validation

    For local_i = LBound(body) To UBound(body)
        Dim entry As Variant
        entry = body(local_i)
        
        If Not IsArray(entry) Then ' Ensure entry is an array
            success = False
            Exit Sub
        End If
        
        success = (UBound(head) = UBound(entry)) And (LBound(head) = LBound(entry))
        If Not success Then Exit For ' Exit loop if dimensions don't match
        
        Dim props As Dictionary
        Set props = New Dictionary
        Dim j As Long ' Use local loop variable
        For j = LBound(head) To UBound(head)
            props(head(j)) = entry(j)
        Next
        pushItem buffer, props
    Next
    
    If success Then ' Only assign data if all entries were successful
        data = buffer
    Else
        data = Array() ' Return empty array on failure
    End If
    
End Sub

Public Sub filterElements(root As Variant, conditions As Variant, inclusive As Boolean, ByRef result As Variant, ByRef success As Boolean)
    Dim data As Variant
    Dim k As Variant
    Dim ret As Variant
    Dim decision As Boolean
    Dim ok As Boolean
    
    success = False ' Initialize
    
    If IsArray(root) Then
        data = Array()
        For k = LBound(root) To UBound(root) ' Use LBound for robustness
            evaluateExpression root(k), conditions, ret, ok
            decision = False
            Select Case False ' This logic seems inverted; if ok is False, or ret is False, then decision is True for inclusive=False
                Case ok
                Case VarType(ret) = vbBoolean
                Case ret ' If ret is True (and ok is True, and VarType is Boolean)
                Case Else ' This means ok=True, VarType=Boolean, ret=False OR ok=False OR VarType<>Boolean
                    decision = True
            End Select
            If decision Xor Not inclusive Then ' (True Xor True = False), (False Xor True = True), (True Xor False = True), (False Xor False = False)
                                            ' If inclusive=True: decision must be False. (ret was True)
                                            ' If inclusive=False: decision must be True. (ret was False or error)
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
                assign root(k), data(k) ' Use assign for object/value
            End If
        Next
        Set result = data
        success = True
    Else
        success = False ' Root is not Array or Dictionary
    End If
End Sub

Public Sub groupElements(root As Variant, path As Variant, full As Boolean, ByRef result As Variant, ByRef success As Boolean)
    Dim k As Variant
    Dim entry As Variant
    Dim currentExists As Boolean ' Renamed from exists to avoid conflict with module var
    
    success = False ' Initialize
    Set result = New Dictionary ' Initialize result as a Dictionary

    If IsArray(root) Then
        Dim buffer As Variant
        buffer = Array() ' Array to hold arrays of items for each group
        Dim groupIndexMap As Dictionary
        Set groupIndexMap = New Dictionary ' Maps group entry value to index in buffer

        For k = LBound(root) To UBound(root) ' Use LBound
            selectElement root(k), path, entry, currentExists
            If currentExists Or full Then
                If Not currentExists Then entry = Null ' Group key for items not matching path (if full=True)
                
                Dim groupKeyStr As String ' Dictionary keys must be strings for Null, numbers etc.
                If IsNull(entry) Then groupKeyStr = "NullGroup" Else groupKeyStr = CStr(entry)

                If Not groupIndexMap.exists(groupKeyStr) Then
                    groupIndexMap(groupKeyStr) = groupIndexMap.count ' Assign next available index
                    pushItem buffer, Array() ' Add a new empty array to buffer for this new group
                End If
                pushItem buffer(groupIndexMap(groupKeyStr)), root(k) ' Add item to its group's array
            End If
        Next
        ' Convert buffer of arrays into result dictionary
        For Each k In groupIndexMap.keys()
            Dim originalGroupKey As Variant
            If k = "NullGroup" Then originalGroupKey = Null Else originalGroupKey = entry ' This is tricky, need original key type
            ' For simplicity, keys in result dictionary will be strings from groupIndexMap
            result(k) = buffer(groupIndexMap(k))
        Next
        success = True
    ElseIf TypeOf root Is Dictionary Then
        For Each k In root.keys()
            selectElement root(k), path, entry, currentExists
            If currentExists Or full Then
                If Not currentExists Then entry = Null
                
                Dim groupKeyStr As String
                If IsNull(entry) Then groupKeyStr = "NullGroup" Else groupKeyStr = CStr(entry)

                If Not result.exists(groupKeyStr) Then
                    Set result(groupKeyStr) = New Dictionary ' Each group is a dictionary
                End If
                assign root(k), result(groupKeyStr)(k) ' Add item to its group's dictionary
            End If
        Next
        success = True
    Else
        success = False ' Root is not Array or Dictionary
    End If
End Sub


Private Sub evaluateExpression(root As Variant, expr As Variant, ByRef result As Variant, ByRef success As Boolean)
    Dim operation As String
    Dim value1 As Variant, value2 As Variant, value3 As Variant
    Dim ok As Boolean
    Dim currentExists As Boolean ' Renamed from exists
    
    success = False ' Default to failure
    result = Empty  ' Default result

    If Not IsArray(expr) Or safeUBound(expr) < 0 Then Exit Sub ' Expression must be a non-empty array
    operation = LCase(CStr(expr(LBound(expr))))

    If Left(operation, 4) = "not " Then
        Dim subexpr As Variant
        subexpr = expr
        subexpr(LBound(expr)) = Mid(operation, 5) ' Get actual operation
        evaluateExpression root, subexpr, value1, ok
        If ok And VarType(value1) = vbBoolean Then
            result = Not value1
            success = True
        End If
        Exit Sub
    End If
    
    Select Case operation
        Case "value"
            If safeUBound(expr) < LBound(expr) + 1 Then Exit Sub ' Path argument missing
            selectElement root, expr(LBound(expr) + 1), value1, currentExists
            success = currentExists
            If success Then assign value1, result
        Case "count"
            If safeUBound(expr) < LBound(expr) + 1 Then Exit Sub
            selectElement root, expr(LBound(expr) + 1), value1, currentExists
            success = currentExists
            If success Then
                If IsArray(value1) Then
                    If safeUBound(value1) = -1 Then result = 0 Else result = UBound(value1) - LBound(value1) + 1
                ElseIf TypeOf value1 Is Dictionary Then
                    result = value1.count
                Else
                    success = False ' Cannot count non-collection
                End If
            End If
        Case "exists"
            If safeUBound(expr) < LBound(expr) + 1 Then Exit Sub
            selectElement root, expr(LBound(expr) + 1), value1, currentExists
            result = currentExists
            success = True
        Case "=", "<>", ">", ">=", "<", "<="
            If safeUBound(expr) < LBound(expr) + 2 Then Exit Sub ' Needs two operands
            If isScalar(expr(LBound(expr) + 1)) Then value1 = expr(LBound(expr) + 1) Else evaluateExpression root, expr(LBound(expr) + 1), value1, ok
            If Not ok And Not isScalar(expr(LBound(expr) + 1)) Then Exit Sub
            If Not isScalar(value1) Then Exit Sub ' Must resolve to scalar

            If isScalar(expr(LBound(expr) + 2)) Then value2 = expr(LBound(expr) + 2) Else evaluateExpression root, expr(LBound(expr) + 2), value2, ok
            If Not ok And Not isScalar(expr(LBound(expr) + 2)) Then Exit Sub
            If Not isScalar(value2) Then Exit Sub

            Select Case operation
                Case "=": result = (value1 = value2)
                Case "<>": result = (value1 <> value2)
                Case ">": result = (value1 > value2)
                Case ">=": result = (value1 >= value2)
                Case "<": result = (value1 < value2)
                Case "<=": result = (value1 <= value2)
            End Select
            success = True
        Case "[]", "[)", "(]", "()" ' Interval checks
            If safeUBound(expr) < LBound(expr) + 3 Then Exit Sub ' Needs value and two bounds
            If isScalar(expr(LBound(expr) + 1)) Then value1 = expr(LBound(expr) + 1) Else evaluateExpression root, expr(LBound(expr) + 1), value1, ok
            If Not ok And Not isScalar(expr(LBound(expr) + 1)) Then Exit Sub
            If Not isScalar(value1) Then Exit Sub

            If isScalar(expr(LBound(expr) + 2)) Then value2 = expr(LBound(expr) + 2) Else evaluateExpression root, expr(LBound(expr) + 2), value2, ok
            If Not ok And Not isScalar(expr(LBound(expr) + 2)) Then Exit Sub
            If Not isScalar(value2) Then Exit Sub
            
            If isScalar(expr(LBound(expr) + 3)) Then value3 = expr(LBound(expr) + 3) Else evaluateExpression root, expr(LBound(expr) + 3), value3, ok
            If Not ok And Not isScalar(expr(LBound(expr) + 3)) Then Exit Sub
            If Not isScalar(value3) Then Exit Sub

            Select Case operation
                Case "[]": result = (value1 >= value2 And value1 <= value3)
                Case "[)": result = (value1 >= value2 And value1 < value3)
                Case "(]": result = (value1 > value2 And value1 <= value3)
                Case "()": result = (value1 > value2 And value1 < value3)
            End Select
            success = True
        Case "or", "and" ' Handle multiple expressions for OR and AND
            Dim idx As Long
            result = (operation = "and") ' Initial value for AND (True) / OR (False)
            For idx = LBound(expr) + 1 To UBound(expr)
                evaluateExpression root, expr(idx), value1, ok
                If Not ok Or VarType(value1) <> vbBoolean Then Exit Sub ' All expressions must evaluate to boolean
                If operation = "or" Then
                    If value1 Then result = True: Exit For ' Short-circuit OR
                Else ' AND
                    If Not value1 Then result = False: Exit For ' Short-circuit AND
                End If
            Next
            success = True
        Case "xor" ' XOR typically takes two arguments
            If safeUBound(expr) < LBound(expr) + 2 Then Exit Sub
            evaluateExpression root, expr(LBound(expr) + 1), value1, ok
            If Not ok Or VarType(value1) <> vbBoolean Then Exit Sub
            evaluateExpression root, expr(LBound(expr) + 2), value2, ok
            If Not ok Or VarType(value2) <> vbBoolean Then Exit Sub
            result = (value1 Xor value2)
            success = True
    End Select
End Sub


Public Sub sort(root As Variant, path As Variant, ascending As Boolean, ByRef result As Variant)
    ascend = ascending ' Module-level variable
    Dim sample() As Variant
    Dim index() As Variant
    Dim data As Variant
    Dim last As Long
    Dim k As Long, local_i As Long ' local_i to avoid conflict with module 'i'
    Dim entry As Variant
    Dim currentExists As Boolean

    If IsArray(root) Then
        last = safeUBound(root)
        If last >= LBound(root) Then ' Check if array has elements
            ReDim sample(LBound(root) To last)
            ReDim index(LBound(root) To last)
            ReDim data(LBound(root) To last)
            For k = LBound(root) To last
                index(k) = k ' Store original index
                sample(k) = Null ' Default sort value
                selectElement root(k), path, entry, currentExists
                If currentExists And Not IsEmpty(entry) And isScalar(entry) Then
                    sample(k) = entry
                End If
            Next
            quickSortIndex sample, index ' Sorts 'index' array based on 'sample' values
            For k = LBound(root) To last
                local_i = index(k) ' Get original index from sorted index array
                assign root(local_i), data(k) ' Use assign
            Next
        Else
            data = Array() ' Empty array if root is empty
        End If
        result = data
    ElseIf TypeOf root Is Dictionary Then
        Set data = New Dictionary
        Dim keys As Variant
        keys = root.keys()
        last = UBound(keys)
        If last >= 0 Then
            ReDim sample(0 To last)
            ReDim index(0 To last)
            Dim keyStr As String
            For k = 0 To last
                keyStr = keys(k)
                index(k) = k ' Store index of key in keys array
                sample(k) = Null
                selectElement root(keyStr), path, entry, currentExists
                If currentExists And Not IsEmpty(entry) And isScalar(entry) Then
                    sample(k) = entry
                End If
            Next
            quickSortIndex sample, index
            For k = 0 To last
                local_i = index(k)
                keyStr = keys(local_i)
                assign root(keyStr), data(keyStr) ' Use assign
            Next
        End If
        Set result = data
    Else
        assign root, result ' Non-collection, return as is
    End If
End Sub


Private Sub quickSortIndex(sample() As Variant, index() As Variant)
    Dim last As Long, first As Long
    first = LBound(index)
    last = UBound(index)

    If last > first Then ' Ensure there's more than one element to sort
        Dim ltArray() As Variant, eqArray() As Variant, gtArray() As Variant
        ltArray = Array(): eqArray = Array(): gtArray = Array()
        
        Dim p As Long
        p = index(first + Int((last - first + 1) / 2)) ' Choose pivot from index array, get corresponding sample value
        Dim pivotValue As Variant
        pivotValue = sample(p) ' The value from sample array to pivot around

        Dim currentIdxVal As Long
        Dim elt As Variant
        Dim local_i As Long

        For local_i = first To last
            currentIdxVal = index(local_i)
            elt = sample(currentIdxVal) ' Value from sample array for current index

            Dim comparisonResult As Integer ' -1 for lt, 0 for eq, 1 for gt
            If IsNull(pivotValue) And IsNull(elt) Then comparisonResult = 0
            ElseIf IsNull(pivotValue) Then comparisonResult = 1 ' Nulls are greater
            ElseIf IsNull(elt) Then comparisonResult = -1    ' Nulls are greater
            ElseIf elt < pivotValue Then comparisonResult = -1
            ElseIf elt > pivotValue Then comparisonResult = 1
            Else comparisonResult = 0
            End If
            
            If Not ascend Then comparisonResult = comparisonResult * -1 ' Invert comparison for descending

            If comparisonResult = -1 Then
                pushItem ltArray, currentIdxVal
            ElseIf comparisonResult = 1 Then
                pushItem gtArray, currentIdxVal
            Else
                pushItem eqArray, currentIdxVal
            End If
        Next
        
        quickSortIndex sample, ltArray
        quickSortIndex sample, gtArray
        
        p = first ' Reset p to be the current position in the main index array
        If safeUBound(ltArray) >= LBound(ltArray) Then
            For local_i = LBound(ltArray) To UBound(ltArray): index(p) = ltArray(local_i): p = p + 1: Next
        End If
        If safeUBound(eqArray) >= LBound(eqArray) Then
            For local_i = LBound(eqArray) To UBound(eqArray): index(p) = eqArray(local_i): p = p + 1: Next
        End If
        If safeUBound(gtArray) >= LBound(gtArray) Then
            For local_i = LBound(gtArray) To UBound(gtArray): index(p) = gtArray(local_i): p = p + 1: Next
        End If
    End If
End Sub

Public Sub selectElement(root As Variant, path As Variant, ByRef entry As Variant, ByRef exists As Boolean)
    Dim elt As Variant
    Dim i As Long 
    Dim localPathIsString As Boolean
    Dim parsedPathParts() As Variant 

    exists = False
    entry = Empty 

    If Not IsArray(path) Then
        localPathIsString = True
        If VarType(path) <> vbString Then
            Exit Sub 
        End If

        Dim pathString As String
        pathString = CStr(path)

        If pathString = "" Then
            assign root, entry
            exists = True
            Exit Sub
        Else
            Dim elts As Variant
            elts = Split(Replace(Replace(Replace(pathString, ".", "|."), "[", "|["), "|.|[", "|.["), "|")
            
            If safeUBound(elts) < 0 Then
                 Exit Sub 
            End If

            If elts(LBound(elts)) <> "" Then
                Exit Sub 
            End If

            If UBound(elts) = LBound(elts) And elts(LBound(elts)) = "" Then 
                assign root, entry
                exists = True
                Exit Sub
            End If
            
            Dim tempParts() As Variant
            ReDim tempParts(LBound(elts) To UBound(elts)) ' Initial conservative sizing
            Dim pathPartCount As Long
            pathPartCount = 0

            For i = LBound(elts) + 1 To UBound(elts) 
                elt = elts(i)
                If Left(elt, 1) = "." Then
                    tempParts(pathPartCount) = Mid(elt, 2)
                    pathPartCount = pathPartCount + 1
                ElseIf Left(elt, 1) = "[" And Right(elt, 1) = "]" Then
                    elt = Mid(elt, 2, Len(elt) - 2)
                    If IsNumeric(elt) Then
                        tempParts(pathPartCount) = CLng(elt)
                        pathPartCount = pathPartCount + 1
                    Else
                        Exit Sub 
                    End If
                ElseIf elt <> "" Then 
                    Exit Sub 
                ElseIf elt = "" And i = UBound(elts) Then
                    ' Handled by pathPartCount not incrementing for trailing empty segment
                ElseIf elt = "" And i < UBound(elts) Then
                     Exit Sub 
                End If
            Next

            If pathPartCount > 0 Then
                ReDim parsedPathParts(pathPartCount - 1)
                For i = 0 To pathPartCount - 1
                    parsedPathParts(i) = tempParts(i)
                Next
            Else 
                 If pathString <> "" And UBound(elts) = LBound(elts) And elts(LBound(elts)) = "" Then
                    ' Path was just "." or "[]", handled above by assigning root
                 ElseIf pathString <> "" And UBound(elts) > LBound(elts) And pathPartCount = 0 Then
                    ' Path like ".prop1." or ".prop1[]" - select 'prop1'
                    ' This means elts(1) was valid, pathPartCount became 1, then elts(2) was empty.
                    ' The logic needs to handle this if the intent is to select the parent of the final empty segment.
                    ' For now, this will likely result in exists=False due to empty parsedPathParts or failed traversal.
                    ' Re-evaluating this specific case: if path is "obj.", parsing makes parsedPathParts("obj"). Loop below handles it.
                    ' If path is "obj..", parsing makes parsedPathParts("obj", ""). Loop below handles it.
                    Exit Sub ' No valid parts to traverse
                 Else
                    Exit Sub
                 End If
            End If
        End If
    Else
        localPathIsString = False
        If safeUBound(path) < LBound(path) Then ' Empty array like Array()
             parsedPathParts = Array() 
        Else
             parsedPathParts = path 
        End If

        If Not (LBound(parsedPathParts) <= UBound(parsedPathParts)) Then
            assign root, entry
            exists = True
            Exit Sub
        End If
    End If

    assign root, entry 

    If Not (LBound(parsedPathParts) <= UBound(parsedPathParts)) Then
         If (localPathIsString And pathString = "") Or (Not localPathIsString And safeUBound(path) < LBound(path)) Then
            exists = True ' Root is the entry for explicitly empty paths
            Exit Sub
        Else
            ' This implies a path like "." or "[]" which got parsed to zero usable segments
            ' but wasn't an explicitly empty path string/array.
            ' The root is effectively selected.
            exists = True 
            Exit Sub
        End If
    End If

    For i = LBound(parsedPathParts) To UBound(parsedPathParts)
        elt = parsedPathParts(i)
        
        If VarType(elt) = vbString And CStr(elt) = "" Then
            If i = UBound(parsedPathParts) Then ' Trailing empty segment means select current entry
                exists = True
                Exit For
            Else ' Empty segment in middle of path
                exists = False
                Exit For
            End If
        End If

        If IsArray(entry) Then
            If Not IsNumeric(elt) Then 
                exists = False
                Exit For
            End If
            Dim arrIndex As Long
            On Error Resume Next ' Temporarily for LBound/UBound check on potentially non-array
            arrIndex = CLng(elt)
            If Err.Number <> 0 Then exists = False: Err.Clear: On Error GoTo 0: Exit For
            On Error GoTo 0
            
            Dim lb As Long, ub As Long
            lb = LBound(entry)
            ub = UBound(entry)
            If arrIndex < lb Or arrIndex > ub Then 
                exists = False
                Exit For
            End If
            assign entry(arrIndex), entry 
        ElseIf TypeOf entry Is Dictionary Then
            If Not entry.exists(elt) Then 
                exists = False
                Exit For
            End If
            assign entry(elt), entry 
        Else 
            exists = False 
            Exit For
        End If
    Next i

    If i > UBound(parsedPathParts) Then
        exists = True
    Else
        ' If exists was True from a trailing empty segment, keep it. Otherwise, it's False.
        If Not (VarType(elt) = vbString And CStr(elt) = "" And i = UBound(parsedPathParts) And exists = True) Then
             exists = False
        End If
    End If
End Sub

Public Sub joinSubDicts(acc As Dictionary, src As Dictionary, Optional addNew As Boolean = True)
    If Not (TypeOf acc Is Dictionary And TypeOf src Is Dictionary) Then Exit Sub
    Dim key As Variant
    For Each key In src.keys()
        If TypeOf src(key) Is Dictionary Then
            Dim srcSubDict As Dictionary
            Set srcSubDict = src(key)
            Dim accSubDict As Dictionary
            Set accSubDict = Nothing
            If acc.exists(key) Then
                If TypeOf acc(key) Is Dictionary Then Set accSubDict = acc(key)
            End If
            If accSubDict Is Nothing Then
               Set accSubDict = New Dictionary
               Set acc(key) = accSubDict
            End If
            joinDicts accSubDict, srcSubDict, addNew
        End If
    Next
End Sub

Public Sub joinDicts(acc As Dictionary, src As Dictionary, Optional addNew As Boolean = True)
    If Not (TypeOf acc Is Dictionary And TypeOf src Is Dictionary) Then Exit Sub
    Dim key As Variant
    If addNew Then
        For Each key In src.keys(): assign src(key), acc(key): Next
    Else
        For Each key In src.keys(): If acc.exists(key) Then assign src(key), acc(key): Next
    End If
End Sub

Public Sub slice(src As Variant, Optional ByRef result As Variant, Optional ByVal a As Variant, Optional ByVal b As Variant)
    Dim m As Long, lb As Long ' Added LBound
    Dim useResultParam As Boolean
    useResultParam = Not IsMissing(result)

    If IsArray(src) Then
        lb = LBound(src)
        m = UBound(src)
    ElseIf TypeOf src Is Dictionary Then
        lb = 0 ' Dictionaries don't have LBound, keys are used; effectively 0-indexed for slicing by count
        m = src.count - 1
    Else ' Not a collection, cannot slice
        If useResultParam Then assign src, result Else ' No-op if result not passed
        Exit Sub
    End If

    If IsMissing(a) Then a = lb
    If IsMissing(b) Then b = m

    Dim temp As Variant
    Dim idx As Long ' Renamed from i
    Dim d As Long
    
    If Not (IsNumeric(a) And IsNumeric(b)) Then
        If useResultParam Then assign src, result Else src = src
        Exit Sub
    End If
    
    a = CLng(a): b = CLng(b) ' Ensure Long

    Dim void As Boolean: Dim full As Boolean
    If (a < lb And b < lb) Or (a > m And b > m) Then void = True
    If a <= lb And b >= m Then full = True
    
    If Not void Then ' Adjust bounds if partially overlapping
        If a < lb Then a = lb
        If a > m Then a = m
        If b < lb Then b = lb
        If b > m Then b = m
    End If

    If IsArray(src) Then
        If void Then temp = Array()
        ElseIf full Then temp = src
        ElseIf a > b Then temp = Array() ' Invalid range for array if not stepping backwards (which this doesn't explicitly)
        Else
            ReDim temp(a To b) ' Direct slice
            For idx = a To b
                assign src(idx), temp(idx)
            Next
        End If
        If useResultParam Then result = temp Else src = temp
    ElseIf TypeOf src Is Dictionary Then
        If void Then Set temp = New Dictionary: temp.CompareMode = src.CompareMode
        ElseIf full Then Set temp = jsonExt.cloneDictionary(src)
        Else
            Set temp = New Dictionary: temp.CompareMode = src.CompareMode
            Dim keys As Variant: keys = src.keys()
            If a > b Or a > UBound(keys) Or b < LBound(keys) Then ' Invalid range for keys
                ' Do nothing, temp remains empty dictionary
            Else
                 ' Adjust a and b to be valid 0-based indices for keys array
                If a < 0 Then a = 0
                If b > UBound(keys) Then b = UBound(keys)

                For idx = a To b ' Iterate through selected key indices
                    If idx >= LBound(keys) And idx <= UBound(keys) Then ' Ensure index is valid for keys array
                        assign src(keys(idx)), temp(keys(idx))
                    End If
                Next
            End If
        End If
        If useResultParam Then Set result = temp Else Set src = temp
    Else ' Should have exited if not array/dictionary
        If useResultParam Then assign src, result
    End If
End Sub


Public Sub getAvg(root As Variant, path As Variant, ByRef avg As Variant, ByRef sum As Variant, ByRef qty As Long)
    Dim k As Variant, entry As Variant, currentExists As Boolean
    sum = 0: qty = 0: avg = 0 ' Initialize

    If IsArray(root) Then
        For k = LBound(root) To UBound(root)
            selectElement root(k), path, entry, currentExists
            If currentExists And IsNumeric(entry) Then qty = qty + 1: sum = sum + CDbl(entry)
        Next
    ElseIf TypeOf root Is Dictionary Then
        For Each k In root.keys()
            selectElement root(k), path, entry, currentExists
            If currentExists And IsNumeric(entry) Then qty = qty + 1: sum = sum + CDbl(entry)
        Next
    End If
    If qty > 0 Then avg = sum / qty
End Sub

Public Sub getMax(root As Variant, path As Variant, ByRef key As Variant, ByRef ret As Variant, ByRef qty As Long)
    Dim currentMax As Variant: currentMax = Null ' Use Null to handle negative numbers correctly
    Dim k As Variant, entry As Variant, currentExists As Boolean, numEntry As Double
    qty = 0: key = Empty: ret = Empty

    If IsArray(root) Then
        For k = LBound(root) To UBound(root)
            selectElement root(k), path, entry, currentExists
            If currentExists And IsNumeric(entry) Then
                numEntry = CDbl(entry): qty = qty + 1
                If IsNull(currentMax) Or numEntry > currentMax Then currentMax = numEntry: key = k
            End If
        Next
    ElseIf TypeOf root Is Dictionary Then
        For Each k In root.keys()
            selectElement root(k), path, entry, currentExists
            If currentExists And IsNumeric(entry) Then
                numEntry = CDbl(entry): qty = qty + 1
                If IsNull(currentMax) Or numEntry > currentMax Then currentMax = numEntry: key = k
            End If
        Next
    End If
    If qty > 0 Then ret = currentMax
End Sub

Public Sub getMin(root As Variant, path As Variant, ByRef key As Variant, ByRef ret As Variant, ByRef qty As Long)
    Dim currentMin As Variant: currentMin = Null
    Dim k As Variant, entry As Variant, currentExists As Boolean, numEntry As Double
    qty = 0: key = Empty: ret = Empty

    If IsArray(root) Then
        For k = LBound(root) To UBound(root)
            selectElement root(k), path, entry, currentExists
            If currentExists And IsNumeric(entry) Then
                numEntry = CDbl(entry): qty = qty + 1
                If IsNull(currentMin) Or numEntry < currentMin Then currentMin = numEntry: key = k
            End If
        Next
    ElseIf TypeOf root Is Dictionary Then
        For Each k In root.keys()
            selectElement root(k), path, entry, currentExists
            If currentExists And IsNumeric(entry) Then
                numEntry = CDbl(entry): qty = qty + 1
                If IsNull(currentMin) Or numEntry < currentMin Then currentMin = numEntry: key = k
            End If
        Next
    End If
    If qty > 0 Then ret = currentMin
End Sub

Function safeUBound(a As Variant) As Long
    safeUBound = -1 ' Default for error or uninitialized array
    If Not IsArray(a) Then Exit Function
    On Error Resume Next
    safeUBound = UBound(a)
    Err.Clear
End Function

Function isScalar(v As Variant) As Boolean
    Select Case VarType(v)
        Case vbEmpty, vbNull, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString, vbBoolean, vbDecimal, vbByte
            isScalar = True
        Case Else
            isScalar = False
    End Select
End Function

Function cloneDictionary(srcDict As Dictionary) As Dictionary
    Dim destDict As Dictionary
    Set destDict = New Dictionary
    If TypeOf srcDict Is Dictionary Then
        destDict.CompareMode = srcDict.CompareMode
        Dim key As Variant
        For Each key In srcDict.keys()
            assign srcDict(key), destDict(key) ' Use assign
        Next
    End If
    Set cloneDictionary = destDict
End Function

Sub deepClone(srcElt As Variant, ByRef destElt As Variant)
    If IsArray(srcElt) Then
        Dim lb As Long, ub As Long, i As Long
        lb = LBound(srcElt): ub = UBound(srcElt)
        ReDim destElt(lb To ub)
        For i = lb To ub
            Dim tmp As Variant
            deepClone srcElt(i), tmp
            assign tmp, destElt(i) ' Use assign
        Next
    ElseIf TypeOf srcElt Is Dictionary Then
        Set destElt = New Dictionary
        destElt.CompareMode = srcElt.CompareMode
        Dim key As Variant, tmp As Variant
        For Each key In srcElt.keys()
            deepClone srcElt(key), tmp
            assign tmp, destElt(key) ' Use assign
        Next
    ElseIf IsObject(srcElt) Then ' For other object types, assign by reference
        Set destElt = srcElt
    Else ' Scalar
        destElt = srcElt
    End If
End Sub

Sub pushItem(ByRef destArray As Variant, sourceElement As Variant, Optional optionAppend As Boolean = True, Optional optionNestArrays As Boolean = True)
    Dim lb As Long, ub As Long
    
    If Not optionAppend Or IsEmpty(destArray) Or VarType(destArray) < vbArray Then
        If IsArray(sourceElement) And Not optionNestArrays Then
            destArray = sourceElement ' Assign directly if not appending and source is array
        Else
            ReDim destArray(0 To 0)
            assign sourceElement, destArray(0)
        End If
        Exit Sub
    End If

    lb = LBound(destArray)
    ub = UBound(destArray)

    If IsArray(sourceElement) And Not optionNestArrays Then
        Dim srcLb As Long, srcUb As Long, srcLen As Long
        srcLb = LBound(sourceElement): srcUb = UBound(sourceElement)
        srcLen = srcUb - srcLb + 1
        If srcLen > 0 Then
            ReDim Preserve destArray(lb To ub + srcLen)
            Dim i As Long, j As Long
            j = ub + 1
            For i = srcLb To srcUb
                assign sourceElement(i), destArray(j)
                j = j + 1
            Next
        End If
    Else
        ReDim Preserve destArray(lb To ub + 1)
        assign sourceElement, destArray(ub + 1)
    End If
End Sub

Sub assign(source As Variant, ByRef dest As Variant)
    If IsObject(source) Then
        Set dest = source
    Else
        dest = source
    End If
End Sub
