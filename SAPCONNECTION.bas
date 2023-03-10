Attribute VB_Name = "SAPCONNECTION"
Public TABMELDUNGFAIL() As Variant, MELDUNGFAIL As String
Public TABMELDUNGFAIL2() As Variant
Public w As Integer
Public SAPGUI
Public WshShell
Public PROC
Public APPL
Public CONNECTION
Public SESSION
Public LANCERSAPPAC As Boolean
Public CODE As String, LOGON


Sub CONNECTIONSAP()



'***************************************************************************************************************
'***************************************************************************************************************
On Error GoTo ERR:
Dim CONNECTER As Boolean
Dim i As Integer
Dim LOGON
'Dim mon_mot_de_passe As String
'Dim mon_TKA As String

Set SAPGUI = GetObject("SAPGUI")

'***************************************************************************************************************
'SI NON CONNECT�
If CONNECTER = True Then
'mon_TKA = Range("B2").value
'mon_mot_de_passe = Range("A2").value

    'session.findById("wnd[0]").maximize
'    SESSION.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "mon_TKA"
'    SESSION.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "mon_mot_de_passe"
'    SESSION.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "EN"
'    SESSION.findById("wnd[0]/usr/txtRSYST-MANDT").SetFocus
'    SESSION.findById("wnd[0]/usr/txtRSYST-MANDT").caretPosition = 0
'    SESSION.findById("wnd[0]").sendVKey 0

    'Lancement automatique SAP
    Set WshShell = CreateObject("WScript.Shell")
    Set PROC = WshShell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
     Application.Wait Now + TimeValue("0:00:05")
    Set SAPGUI = GetObject("SAPGUI")
    Set APPL = SAPGUI.GetScriptingEngine
    Application.Wait Now + TimeValue("0:00:01")
LANCERSAPPACZQRS:
    Set CONNECTION = APPL.Openconnection("DAc: ERP Produktion [PAC]", True)
    'SendKeys "%O"
    Set SESSION = CONNECTION.Children(0)
    'SESSION.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "EN"
    'LOGON = InputBox("Entrez votre logon (TKA0xxxxx...) :", "Please enter your SAP DAC Logon")
    'LOGON = ID.mon_TKA
    LOGON = mon_TKA
    If Left(LOGON, 3) = "TKA" _
    Or Left(LOGON, 2) = "NG" Then
    
        SESSION.findById("wnd[0]/usr/txtRSYST-BNAME").Text = LOGON
        
    ElseIf Left(LOGON, 3) = "220" _
    Or Left(LOGON, 3) = "210" Then
    
        SESSION.findById("wnd[0]/usr/txtRSYST-BNAME").Text = CHERCHERSAPLOGON
    
    End If
    
    'InputBox("Please enter your Login :", "Login")
'    SESSION.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = InputBox("Entrez votre mot de passe :", "Please enter your SAP DAC Password")
    'SESSION.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = ID.mon_mot_de_passe
    SESSION.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = mon_mot_de_passe
    SESSION.findById("wnd[0]").sendVKey 0
    SESSION.findById("wnd[0]").maximize
    SESSION.findById("wnd[0]").SetFocus
    'Si un SESSION est d�j� ouverte, valider puis ouvrir
    If SESSION.Children.Count > 1 Then
        SESSION.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select
        SESSION.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").SetFocus
        SESSION.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
'***************************************************************************************************************
'SI D�J� CONNECT�

ElseIf CONNECTER = False Then
    
    Set SAPGUI = GetObject("SAPGUI")
    Set APPL = SAPGUI.GetScriptingEngine
    Set CONNECTION = APPL.Children(0)
    Set SESSION = CONNECTION.Children(0)

    'While Connection.Children.Count > 1
        'Set SESSION = Connection.Children(0)
        'SESSION.findById("wnd[0]").Close
    'Wend
   
    'Set SESSION = Connection.Children(0)
    SESSION.CreateSESSION
    Application.Wait Now + TimeValue("0:00:02")
    'SESSION.findById("wnd[0]").Close
    For i = 1 To CONNECTION.Children.Count
        If CONNECTION.Children(CONNECTION.Children.Count - i).info.transaction = "SESSION_MANAGER" Then
        
            Set SESSION = CONNECTION.Children(CONNECTION.Children.Count - i)
            Exit For
            
        End If
    Next

    
End If
'***************************************************************************************************************
'***************************************************************************************************************
Exit Sub

ERR:
If ERR.Number = "-2147221020" Then

    CONNECTER = True
    Resume Next
    
ElseIf ERR.Number = "614" Then

    LANCERSAPPAC = True
    GoTo LANCERSAPPACZQRS
End If
End Sub


