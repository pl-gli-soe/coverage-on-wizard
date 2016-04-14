Attribute VB_Name = "MainModule"
Public Sub main(ich As InitConfigHandler)
    
    ' MAIN SECTION
    ' ---------------------------------------------------------------------
    Dim krowa As CowHandler
    Set krowa = New CowHandler
    
    With krowa
        .init ich
    End With
    
    Set krowa = Nothing
    
    ' ---------------------------------------------------------------------
End Sub
