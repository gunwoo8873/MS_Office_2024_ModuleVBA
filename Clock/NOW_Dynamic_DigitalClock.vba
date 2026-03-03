Sub Current_Digital_Clock()
    Range("A1").Value = Now
    Application.OnTime Now + TimeValue("00:00:01"), "current_digital_clock"
End Sub