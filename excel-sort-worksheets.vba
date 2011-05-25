Function Partition(B_I As Integer, E_I As Integer)
    Dim Pivot As String
    Pivot = ActiveWorkbook.Worksheets(B_I).Name
    
    Do While (True)
        Do While (StrComp(ActiveWorkbook.Worksheets(B_I).Name, Pivot) < 0 And B_I < ActiveWorkbook.Worksheets.Count)
            B_I = B_I + 1
        Loop
        Do While (StrComp(ActiveWorkbook.Worksheets(E_I).Name, Pivot) > 0 And E_I > 1)
            E_I = E_I - 1
        Loop
        
        If (ActiveWorkbook.Worksheets(B_I).Name > ActiveWorkbook.Worksheets(E_I).Name) Then
            ActiveWorkbook.Worksheets(E_I).Move _
                after:=ActiveWorkbook.Worksheets(B_I)
            ActiveWorkbook.Worksheets(B_I).Move _
                after:=ActiveWorkbook.Worksheets(E_I)
        Else
            Exit Do
        End If
    Loop

    Partition = E_I
End Function

Sub SortWorksheetsHelper(B_I As Integer, E_I As Integer)
    If (B_I < E_I) Then
        ' Everything in VBA is passed by reference unless you make a copy
        Dim B_I_Copy As Integer
        B_I_Copy = B_I
        Dim E_I_Copy As Integer
        E_I_Copy = E_I

        Dim PivotIndex As Integer
        PivotIndex = Partition(B_I, E_I)

        Call SortWorksheetsHelper(B_I, PivotIndex - 1)
        Call SortWorksheetsHelper(PivotIndex + 1, E_I)
    End If
End Sub

Sub SortWorksheets()
    Call SortWorksheetsHelper(1, ActiveWorkbook.Worksheets.Count)
End Sub
