Attribute VB_Name = "ID"

Public mon_TKA As String
Public mon_mot_de_passe As String
Public Langue As String
Public Tools As String

Sub ID()

'Lire les valeurs des valeurs dans les cellules depuis le fichier Excel
mon_mot_de_passe = Range("A2").value
mon_TKA = Range("B2").value
Tools = Range("C2").value
Langue = Range("D2")

End Sub
