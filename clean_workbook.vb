Option Explicit

Public Function DoesItemExist(mySet As Collection, myCheck As String) As Boolean

    Dim elm As Variant
    DoesItemExist = False
    For Each elm In mySet
        If myCheck = elm Then
            DoesItemExist = True
            Exit Function
        End If
    Next
End Function




Sub clean(level As Integer, symbols As Variant, direction As String, Optional sheet As Worksheet)

    Dim wb As Worksheet
    If sheet Is Nothing Then
        Set wb = ActiveSheet
    Else
        Set wb = sheet
    End If
    
    
    'Depending on input datatype, create collection
    Dim s As Collection
        If TypeOf symbols Is Collection Then
           Set s = symbols
        Else
            Set s = New Collection
            s.Add (symbols)
        End If
    
    Dim ur As Range
    Dim rowcount, colcount As Integer
    Set ur = wb.UsedRange
    colcount = ur.Columns.Count
    rowcount = ur.Rows.Count
    
    Dim i As Integer
    If direction = "rows" Then
        For i = rowcount To level + 1 Step (-1)
            If Not DoesItemExist(s, Rows.Rows(i).Cells(level).Value) Then Rows.Rows(i).EntireRow.Delete
        Next i
    ElseIf direction = "columns" Then
        For i = colcount To level + 1 Step (-1)
            If Not DoesItemExist(s, Columns.Columns(i).Cells(level).Value) Then Columns.Columns(i).EntireColumn.Delete
        Next i
    End If
    
    
End Sub


Sub main()
    Dim keepcoll As Collection
    Set keepcoll = New Collection
    keepcoll.Add ("x")
    keepcoll.Add ("a")
    Call clean(1, keepcoll, "columns")
End Sub


