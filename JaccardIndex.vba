    Sub CalculateJaccardIndex()
        Dim ws As Worksheet
        Dim lastRow As Long
        Dim i As Long
        Dim revisedText As String
        Dim coalescedText As String
        Dim jacard As Single
        Dim imax As Long
        

        Let i = 2
        Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name
            imax = getLastRow(4, ws)
        Do
            revisedText = ws.Cells(i, 5).Value ' 5->Column E for "RevisedText"
            coalescedText = ws.Cells(i, 4).Value ' 4->Column D for "CoalescedText"
            jacard = JaccardIndex(revisedText, coalescedText)
            ws.Cells(i, 6).Value = jacard	' 6-> Column F for Jaccard Index
            i = i + 1
        Loop While (i <= imax)
    
    End Sub
    
    Function JaccardIndex(str1 As String, str2 As String) As Double
	'Calculate Jaccard Index as |A intersection B| / |A union B| on distinct elements

        Dim set1 As Collection
        Dim unionSet As Collection
        Dim intersectionSet As Collection
        Dim element As Variant
        
        Set set1 = New Collection
        Set unionSet = New Collection
        Set intersectionSet = New Collection
        

        'Find union
        For Each element In Split(str1, " ")
            set1.Add element
            If Not IsInCollection(element, unionSet) Then
                unionSet.Add element
            End If
        Next element

	'Fin intersection
        For Each element In Split(str2, " ")
            If Not IsInCollection(element, unionSet) Then
                unionSet.Add element
            End If
            If IsInCollection(element, set1) Then
                If Not IsInCollection(element, intersectionSet) Then
                    intersectionSet.Add element
                End If
            End If
        Next element
    

    If unionSet.Count = 0 Then
        JaccardIndex = 0#
    Else
        JaccardIndex = intersectionSet.Count / unionSet.Count * 100#
    End If
    
End Function
    
Function IsInCollection(sitem As Variant, coll As Object)
' Check if sitem is in collection coll
    Dim lcount As Integer
    IsInCollection = False
    For lcount = 1 To coll.Count
        If sitem = coll.Item(lcount) Then
              IsInCollection = True
              Exit Function
        End If
    Next lcount

End Function

Function getLastRow(col As Integer, Optional ws As Worksheet) As Long
' Find last used row in column col
'from: https://stackoverflow.com/questions/38882321/better-way-to-find-last-used-row'

    If ws Is Nothing Then Set ws = ActiveSheet
    
    If ws.Cells(ws.Rows.Count, col).Value <> "" Then
        getLastRow = ws.Cells(ws.Rows.Count, col).Row
        Exit Function
    End If

    getLastRow = ws.Cells(Rows.Count, col).End(xlUp).Row

    If shtRowCount = 1 Then
        If ws.Cells(1, col) = "" Then
            getLastRow = 0
        Else
            getLastRow = 1
        End If
    End If

End Function
