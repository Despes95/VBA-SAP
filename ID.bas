Attribute VB_Name = "ID"

Public mon_TKA As String
Public mon_mot_de_passe As String
Public Langue As String
Public Tools As String

Sub ID()

'Lire les valeurs des valeurs dans les cellules depuis le fichier Excel
mon_TKA = Range("C6").value
mon_mot_de_passe = Range("E6").value
Langue = Range("C9")
Tools = Range("E9").value

End Sub
