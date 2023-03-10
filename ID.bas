Attribute VB_Name = "ID"

Public mon_TKA As String
Public mon_mot_de_passe As String

Sub ID()

'Lire les valeurs de mon_TKA et mon_mot_de_passe depuis le fichier Excel
mon_mot_de_passe = Range("A2").value
mon_TKA = Range("B2").value

End Sub
