Attribute VB_Name = "Admin"
Sub admin()

    ' Admin button so that user doesnt have access to config and org data

    Dim InputPass As String, Password As String
    
    Password = "admin"
    InputPass = VBA.InputBox("Enter Password to toggle access")
    If InputPass = "" Then Exit Sub
    If InputPass <> Password Then
        MsgBox "The Password is incorrect!", vbCritical, "Incorrect Password"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    If Sheets("Bugs_Updates").Visible = xlSheetVisible Then
        Sheets("Bugs_Updates").Visible = xlSheetVeryHidden
        Sheets("ZSET").Visible = xlSheetVeryHidden
        Sheets("ZGB100").Visible = xlSheetVeryHidden
        Sheets("ZZSERVICE").Visible = xlSheetVeryHidden
        Sheets("hh").Visible = xlSheetVeryHidden
        Sheets("ii").Visible = xlSheetVeryHidden
        Sheets("Lists").Visible = xlSheetVeryHidden
        Sheets("OrgData").Visible = xlSheetVeryHidden
        Sheets("DE_CO_EQ").Visible = xlSheetVeryHidden
    Else
        Sheets("Bugs_Updates").Visible = xlSheetVisible
        Sheets("ZSET").Visible = xlSheetVisible
        Sheets("ZGB100").Visible = xlSheetVisible
        Sheets("ZZSERVICE").Visible = xlSheetVisible
        Sheets("hh").Visible = xlSheetVisible
        Sheets("ii").Visible = xlSheetVisible
        Sheets("Lists").Visible = xlSheetVisible
        Sheets("OrgData").Visible = xlSheetVisible
        Sheets("DE_CO_EQ").Visible = xlSheetVisible
    End If
    
    Sheets("Source").Activate
    
    Application.ScreenUpdating = True


End Sub

Sub save_and_create_ticket()
    
    ' saves file in same path as workbook and takes user to SharePoint is task mode
    
    Call EntryPoint
    
    Dim Country As String
    Dim UserInput As String
    
    ' check if its not empty
    If ShSource.Range("A2").Value = "" Then
        MsgBox "No existing data to process!", vbCritical, "Error!"
        Call EntryPoint
        Exit Sub
    End If
    
    UserInput = MsgBox("This will save a copy of this file in the same directory and will redirect you to Aldi South SharePoint for ticket creation.", vbOKCancel)
    
    If UserInput = vbCancel Then
    
        MsgBox "Process stopped.", vbCritical
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
    Else
    End If
    
    Country = Sheets("Source").Range("A2").Value
    
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Aldi_South_New_Stores_Tool_" & Country
    
    On Error Resume Next
    
    Set objshell = CreateObject("Wscript.Shell")
    objshell.Run ("SharePoint Link")
    
    Call ExitPoint

End Sub

Sub EntryPoint()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .StatusBar = "Macro is running, please wait..."
    End With
End Sub
Sub ExitPoint()
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .StatusBar = ""
    End With
End Sub

Sub clear_data()

    Dim UserInput As String

    Call EntryPoint
    
    UserInput = MsgBox("Are you sure you want to delete all data?", vbOKCancel + vbInformation)
    
    If UserInput = vbCancel Then
        MsgBox "Nothing was deleted"
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
    Else
    End If
    
    ShSource.Range("A2:B50").ClearContents
    ShSource.Range("E2:E50").ClearContents
    ShSource.Range("G2:J50").ClearContents
    ShSource.Range("L2:O50").ClearContents
    ShTicket.Range("A2:BB10000").ClearContents
    ShZZservice.Range("A2:BV10000").ClearContents
    ShHeader.Range("A2:AH10000").ClearContents
    ShItem.Range("A2:X10000").ClearContents
    
    MsgBox "Data cleared from tabs: Source & Ticket", vbInformation
    ShSource.Activate
    ShSource.Range("A2").Select
    
    Call ExitPoint

End Sub

