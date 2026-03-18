Private Sub Worksheet_Change(ByVal Target As Range)
    Dim WatchCell As Range
    Dim TimeCell As Range
    Dim StatusCell As Range
    
    Set WatchCell = Me.Range("B2")
    Set TimeCell = Me.Range("C2")
    Set StatusCell = Me.Range("D2")

    If Not Intersect(Target, WatchCell) Is Nothing Then
        Application.EnableEvents = False
        
        TimeCell.Value = Now()

        If DateValue(TimeCell.Value) = Date Then
            StatusCell.Value = "Update"
        Else
            StatusCell.Value = "Fix"
        End If
        
        Application.EnableEvents = True
    End If
End Sub
