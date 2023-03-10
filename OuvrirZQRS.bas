Attribute VB_Name = "OuvrirZQRS"
Sub LANCEMENTSAP()

'*********

'Lancer transaction ZQRS
SESSION.findById("wnd[0]/tbar[0]/okcd").Text = "ZQRS"
SESSION.findById("wnd[0]").sendVKey 0

'Choix de la variante
SESSION.findById("wnd[0]/tbar[1]/btn[17]").press
SESSION.findById("wnd[1]/usr/txtV-LOW").Text = "LMCPROD"
SESSION.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
SESSION.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
SESSION.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
SESSION.findById("wnd[1]/tbar[0]/btn[8]").press

'KDA
SESSION.findById("wnd[0]/usr/btn%_S_KDAUF_%_APP_%-VALU_PUSH").press
SESSION.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "*-49"
SESSION.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "*-48"
SESSION.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "*-42"
'SESSION.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").SetFocus
'SESSION.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 7
SESSION.findById("wnd[1]/tbar[0]/btn[8]").press

'Open meldung
SESSION.findById("wnd[0]/usr/radP_OQMSM").Select
'First open task
SESSION.findById("wnd[0]/usr/radP_FTASK").Select


'Dates Zieltermin
SESSION.findById("wnd[0]/usr/ctxtS_LTRMN-LOW").Text = "01.01.1900"
SESSION.findById("wnd[0]/usr/ctxtS_LTRMN-HIGH").Text = Format(Day(Date), "00") & "." & Format(Month(Date), "00") & "." & Year(Date)

'Layout
SESSION.findById("wnd[0]/usr/ctxtP_VARI").Text = "PRIOLMCNICO"
SESSION.findById("wnd[0]/usr/ctxtP_VARI").SetFocus
SESSION.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 5
SESSION.findById("wnd[0]").sendVKey 0
SESSION.findById("wnd[0]").sendVKey 8

'Filtre Pour les Prio
SESSION.findById("wnd[0]").maximize
SESSION.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, "PRIOK"
SESSION.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "PRIOK"
SESSION.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").ContextMenu
SESSION.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectContextMenuItem "&FILTER"
SESSION.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "1"
SESSION.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").Text = "2"
SESSION.findById("wnd[1]/tbar[0]/btn[0]").press

'SESSION.CreateSESSION
'
'
''Retrouver une nouvelle fenetre ouverte
'Application.Wait Now + TimeValue("0:00:02")
''Session.findById("wnd[0]").Close
'For i = 1 To CONNECTION.Children.Count
'    If CONNECTION.Children(CONNECTION.Children.Count - i).info.transaction = "SESSION_MANAGER" Then
'
'        Set SESSION = CONNECTION.Children(CONNECTION.Children.Count - i)
'        Exit For
'
'    End If
'Next
'
'SESSION.findById("wnd[0]").SetFocus
'SESSION.findById("wnd[0]/tbar[0]/okcd").Text = "QM02"
'SESSION.findById("wnd[0]").sendVKey 0
'SESSION.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").Text = "700438579"
'SESSION.findById("wnd[0]").sendVKey 0
'SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11").Select
'SESSION.findById("wnd[0]").sendVKey 83
'
'Dim value As String
'value = SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-QSMNUM[0,0]").Text
'
'Dim newValue As Integer
'newValue = CInt(value) + 1
'
'SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-QSMNUM[0,1]").Text = CStr(newValue)
'SESSION.findById("wnd[0]").maximize
'SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1,1]").Text = "ZZR"
'SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNCOD[2,1]").Text = "353"
'SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-MATXT[4,1]").Text = Format(Date, "dd.mm.yy dddd")
'SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/cmbVIQMSM-PARVW[8,1]").Key = "VU"
'SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-PARNR[9,1]").Text = "TKA002831"
''session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER").getAbsoluteRow(67).Selected = True
''session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-QSMNUM[0,1]").SetFocus
''session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-QSMNUM[0,1]").caretPosition = 0
''session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/btnFC_ERLEDIGT").press
'
'SESSION.findById("wnd[0]").maximize


    
End Sub


