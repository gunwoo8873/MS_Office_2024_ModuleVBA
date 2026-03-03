Sub Worksheet_Change(ByVal Target As Range)
    Dim WatchRange As Range
    Dim TimeCell As Range
    
    Set WatchRange = Me.Range("B:D")

    If Not Intersect(Target, WatchRange) Is Nothing Then
        Dim rng As Range
        For Each rng In Target
            If rng.Value <> "" And Me.Cells(rng.Row, "E").Value = "" Then
                Me.Cells(rng.Row, "E").Value = Format(Now, "m/d/yyyy hh:mm:ss")
            End If
        Next rng
    End If
End Sub