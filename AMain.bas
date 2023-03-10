Attribute VB_Name = "AMain"
Sub MAINCHEEK()

    Call SAPCONNECTION.CONNECTIONSAP
    Call Cheek.Cheek
    Call ID.ID
    
End Sub

Sub MAINOUVRRIZQRS()

    Call SAPCONNECTION.CONNECTIONSAP
    Call OuvrirZQRS.LANCEMENTSAP
    Call ID.ID

End Sub
