Attribute VB_Name = "InsertCrossRefToTable"
Sub InsertCrossRefToTable()
    Dim fld As Field
    Dim strRefID As String
    Dim i As Long
    
    ' Get the selected table number
    Dim tblNumber As String
    tblNumber = InputBox("Please enter the number of the table to cross-reference:", "Table Number")
    tblNumber = Trim(tblNumber)
    Debug.Print tblNumber
    ' Search for the specified table number in the active fields
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldSequence Then
            If fld.Code.Text Like "*Table*" Then
                    If InStr(1, fld.Result.Text, tblNumber, vbTextCompare) > 0 Then
                        strRefID = fld.Result.Text
                        Selection.InsertCrossReference ReferenceType:="Table", ReferenceKind:= _
                        wdOnlyLabelAndNumber, ReferenceItem:=strRefID, InsertAsHyperlink:=True, _
                        IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
                        Exit For
                    End If
                End If
            End If
    Next fld
    
    ' Display an error message if the specified table number is not found
    If strRefID = "" Then
        MsgBox "Table " & tblNumber & " was not found.", vbExclamation
    End If
End Sub




