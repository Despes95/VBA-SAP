Attribute VB_Name = "ID"

Public mon_TKA As String
Public mon_mot_de_passe As String
Public Langue As String
Public Tools As String

Sub ID()

'Lire les valeurs des valeurs dans les cellules depuis le fichier Excel
mon_TKA = Range("A2").value
mon_mot_de_passe = Range("B2").value
Langue = Range("C2")
Tools = Range("D2").value

End Sub
