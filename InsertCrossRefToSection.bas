Attribute VB_Name = "InsertCrossRefToSection"
Sub InsertCrossRefToSection()
    Dim myHeadings As Variant
    myHeadings = ActiveDocument.GetCrossReferenceItems(wdRefTypeNumberedItem)
    ' Get the selected section number
    Dim secNumber As String
    secNumber = InputBox("Enter section number:", "Section Number")
    Dim i As Long
    For i = LBound(myHeadings) To UBound(myHeadings)
        Dim lastChar As String
        lastChar = Right(myHeadings(i), 1)
        If IsLetter(lastChar) Then
            Dim item As String
            item = myHeadings(i)
            Dim item1 As String
            Dim length As Long
            length = Len(item)
            Dim j As Long
            For j = length To 1 Step -1
                If Mid(item, j, 1) = "." Then
                item1 = Left(item, j - 1)
                Debug.Print item1
                Debug.Print InStr(1, item1, secNumber, vbTextCompare)
                If InStr(1, item1, secNumber, vbTextCompare) = Len(secNumber) Then
                    Selection.InsertCrossReference ReferenceType:="Numbered item", _
                    ReferenceKind:=wdNumberFullContext, _
                    ReferenceItem:=CStr(i), _
                    InsertAsHyperlink:=True, _
                    IncludePosition:=False, _
                    SeparateNumbers:=False, _
                    SeparatorString:=" "
                    Exit For
                    End If
                End If
            Next j
        End If
    Next i
    


End Sub
Function IsLetter(char As String) As Boolean
    IsLetter = UCase(char) Like "[A-Z]"
End Function


