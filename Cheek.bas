Attribute VB_Name = "Cheek"
Sub Cheek()

'    SESSION.findById("wnd[0]").SetFocus
    SESSION.findById("wnd[0]/tbar[0]/okcd").Text = "QM02"
    SESSION.findById("wnd[0]").sendVKey 0
    'SESSION.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").Text = Tools
    SESSION.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").Text = "700438579"
    SESSION.findById("wnd[0]").sendVKey 0
    SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11").Select
    SESSION.findById("wnd[0]").sendVKey 83
    
    Dim value As String
    value = SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-QSMNUM[0,0]").Text
    
    Dim newValue As Integer
    newValue = CInt(value) + 1
    
    SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-QSMNUM[0,1]").Text = CStr(newValue)
    SESSION.findById("wnd[0]").maximize
    SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1,1]").Text = "ZZR"
    SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNCOD[2,1]").Text = "353"
    SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-MATXT[4,1]").Text = Format(Date, "dd.mm.yy dddd")
    SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/cmbVIQMSM-PARVW[8,1]").Key = "VU"
    SESSION.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-PARNR[9,1]").Text = "TKA002831"
    'session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER").getAbsoluteRow(67).Selected = True
    'session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-QSMNUM[0,1]").SetFocus
    'session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-QSMNUM[0,1]").caretPosition = 0
    'session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/btnFC_ERLEDIGT").press

  

    


End Sub
