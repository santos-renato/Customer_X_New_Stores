Attribute VB_Name = "SAP_Script"
Sub get_consignee_equipment_from_sap()

    Dim Appl As Object
    Dim Connection As Object
    Dim session As Object
    Dim WshShell As Object
    Dim SapGui As Object
    Dim Answer As String
    Dim Path As String
    Dim ticketLR As Long
    Dim ShTicket As Worksheet
    Dim ShDeEQ As Worksheet
    Dim ShVCUST As Worksheet
    Dim ShIE06 As Worksheet
    Dim i As Long
    Dim macroBook As Workbook
    Dim WB As Workbook
    Dim xlApp As Object
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set macroBook = ActiveWorkbook
    
    Set ShTicket = ThisWorkbook.Sheets("Ticket")
    Set ShDeEQ = ThisWorkbook.Sheets("DE_CO_EQ")
    ticketLR = ShTicket.Range("A" & Rows.Count).End(xlUp).row
    
    Answer = MsgBox("Do you want to get Consignee and Equipment numbers from SAP?", vbExclamation + vbYesNo)
    
    If Answer = vbNo Then
        'ThisWorkbook.Close False
        Exit Sub
    End If
    
    MsgBox "Macro will now open SAP session, please wait until you are informed of process completion", vbExclamation
    
    ShDeEQ.Visible = xlSheetVisible
    
    'populate consignee description and serial numbers in DE_CO_EQ
    
    For i = 2 To ticketLR
        ShDeEQ.Cells(i, 1).Value = ShTicket.Cells(i, 3).Value
        ShDeEQ.Cells(i, 3).Value = ShTicket.Cells(i, 34).Value
    Next i
    
    'Of course change for your file directory
    Shell "C:\XXXXXX\SAP\FrontEnd\SAPgui\saplogon.exe", 4
    Set WshShell = CreateObject("WScript.Shell")
    
    Do Until WshShell.AppActivate("SAP Logon ")
        Application.Wait Now + TimeValue("0:00:01")
    Loop
    
    Set WshShell = Nothing
    
    Set SapGui = GetObject("SAPGUI")
    Set Appl = SapGui.GetScriptingEngine
    Set Connection = Appl.Openconnection("010. ERP Germany                - Login without password", _
        True)
    Set session = Connection.Children(0)
    
    'if You need to pass username and password
    'session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "900"
    'session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "user"
    'session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "password"
    'session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "EN"
    
    If session.Children.Count > 1 Then
    
        'Answer = MsgBox("If you already have opened P05 session please click in SAP 2nd option and then click ok on this pop up message")
    
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").SetFocus
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    '
    '    Exit Sub
    
    End If
    
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]").sendVKey 0 'ENTER
    
    'and there goes your code in SAP
    
    Application.DisplayAlerts = False
    
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvcust"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/btn%_SNAME_%_APP_%-VALU_PUSH").press
    ShDeEQ.Range("A2:A" & ticketLR).Copy
    session.findById("wnd[1]").sendVKey 24
    session.findById("wnd[1]").sendVKey 8
    session.findById("wnd[0]").sendVKey 8
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = ThisWorkbook.Path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "VCUST.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nie06"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/btn%_SERNR_%_APP_%-VALU_PUSH").press
    ShDeEQ.Range("C2:C" & ticketLR).Copy
    session.findById("wnd[1]").sendVKey 24
    session.findById("wnd[1]").sendVKey 8
    session.findById("wnd[0]").sendVKey 8
    session.findById("wnd[0]/tbar[1]/btn[32]").press
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 1
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "1"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[1]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    Application.DisplayAlerts = False
    Workbooks("Worksheet in Basis (1)").SaveAs ThisWorkbook.Path & "\IE06.xlsx"
    Application.DisplayAlerts = True
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nex"
    session.findById("wnd[0]").sendVKey 0
    
    Application.DisplayAlerts = True
    
    Workbooks.Open ThisWorkbook.Path & "\VCUST.xlsx"
    Set ShVCUST = ActiveWorkbook.Sheets(1)
    Workbooks.Open ThisWorkbook.Path & "\IE06.xlsx"
    Set ShIE06 = ActiveWorkbook.Sheets(2)
    
    For i = 2 To ticketLR
        ShDeEQ.Cells(i, 3).Value = "'" & ShDeEQ.Cells(i, 3).Value
    Next i
    
    For i = 2 To ticketLR
       ShDeEQ.Cells(i, 2).Value = Application.WorksheetFunction.XLookup(ShDeEQ.Cells(i, 1).Value, ShVCUST.Range("G:G"), ShVCUST.Range("A:A"))
       ShDeEQ.Cells(i, 4).Value = Application.WorksheetFunction.XLookup(ShDeEQ.Cells(i, 3).Value, ShIE06.Range("A:A"), ShIE06.Range("B:B"))
    Next i
    
    Workbooks("VCUST").Close False
    Workbooks("IE06").Close False
    
    Set xlApp = GetObject(ThisWorkbook.Path & "\VCUST.xlsx").Application
    xlApp.Workbooks(1).Close False
    xlApp.Quit
    
    Kill ThisWorkbook.Path & "\VCUST.xlsx"
    Kill ThisWorkbook.Path & "\IE06.xlsx"
    
    MsgBox "Consignees and Equipment numbers extracted from SAP", vbInformation
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

'ThisWorkbook.Close False
End Sub


