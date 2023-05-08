Attribute VB_Name = "InsertCrossRefToFigure"
Sub InsertCrossRefToFigure()
    Dim fld As Field
    Dim strRefID As String
    Dim i As Long
    
    ' Get the selected figure number
    Dim figNumber As String
    figNumber = InputBox("Please enter the number of the figure to cross-reference:", "Figure Number")
    figNumber = Trim(figNumber)
    Debug.Print figNumber
    ' Search for the specified table number in the active fields
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldSequence Then
            If fld.Code.Text Like "*Figure*" Then
            Debug.Print fld.Result.Text
            Debug.Print InStr(1, fld.Result.Text, figNumber, vbTextCompare)
                If InStr(1, fld.Result.Text, figNumber, vbTextCompare) > 0 Then
                        strRefID = fld.Result.Text
                        Selection.InsertCrossReference ReferenceType:="Figure", ReferenceKind:= _
                        wdOnlyLabelAndNumber, ReferenceItem:=strRefID, InsertAsHyperlink:=True, _
                        IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
                    Exit For
                End If
            End If
        End If
    Next fld
    
    ' Display an error message if the specified table number is not found
    If strRefID = "" Then
        MsgBox "Figure " & figNumber & " was not found.", vbExclamation
    End If
End Sub


