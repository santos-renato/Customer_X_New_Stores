Attribute VB_Name = "Ticket_Processor"
Dim ticketLR As Long
Dim sourceLR As Byte
Dim SavePath As String
Dim NewBook As Workbook
Dim i As Long

Sub zset()

    Call EntryPoint
    
    Dim row As Byte
    Dim Year As Integer, Month As Integer, Day As Integer
    
    ticketLR = ShTicket.Range("A" & Rows.Count).End(xlUp).row
    
    ShZSET.Range("C17:AR100000").ClearContents
    
    ' Check if country is Germany
    If ShSource.Range("A2").Value <> "DE" Then
        MsgBox "ZSET will only work with Germany", vbCritical, "Error!"
        Call ExitPoint
        Exit Sub
    End If
    
    ' order customer alphabetically - will be helpful to later save zset per customer
    ShTicket.Activate
    ShTicket.Range("A1").Select
    Selection.AutoFilter
    ShTicket.AutoFilter.Sort.SortFields.Clear
    ShTicket.AutoFilter.Sort.SortFields.Add2 Key:= _
    Range("A1:A" & ticketLR), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
    With ShTicket.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    
    ShZSET.Activate
    
    ' start loop to fill ZSET
    row = 17
    For i = 2 To ticketLR
        ' if customer(i) = customer(i+1) then fill in template
        If ShTicket.Cells(i, 1).Value = ShTicket.Cells(i + 1, 1).Value Then
            ' Customer
            ShZSET.Cells(6, 4).Value = ShTicket.Cells(i, 1).Value
            ' Customer description
            ShZSET.Cells(7, 4).Value = ShTicket.Cells(i, 2).Value
            ' Customer
            ShZSET.Cells(row, 3).Value = ShTicket.Cells(i, 1).Value
            ' EQ description
            ShZSET.Cells(row, 4).Value = ShTicket.Cells(i, 35).Value
            ' Material number
            ShZSET.Cells(row, 5).Value = ShTicket.Cells(i, 33).Value
            ' Serial number
            ShZSET.Cells(row, 6).Value = ShTicket.Cells(i, 34).Value
            ' Warranty start
            Year = Right(ShTicket.Cells(i, 4).Value, 4)
            Month = Mid(ShTicket.Cells(i, 4).Value, 4, 2)
            Day = Left(ShTicket.Cells(i, 4).Value, 2)
            ShZSET.Cells(row, 30).Value = Format(DateSerial(Year, Month, Day), "dd.mm.yyyy")
            ' Warranty end
            ShZSET.Cells(row, 31).Value = Format(DateSerial(Year + 1, Month, Day), "dd.mm.yyyy")
            ' Standort
            ShZSET.Cells(row, 32).Value = ShTicket.Cells(i, 3).Value
            ' Street
            ShZSET.Cells(row, 33).Value = ShTicket.Cells(i, 10).Value
            ' ZIP
            ShZSET.Cells(row, 34).Value = ShTicket.Cells(i, 14).Value
            ' City
            ShZSET.Cells(row, 35).Value = ShTicket.Cells(i, 11).Value
            ' TAG
            ShZSET.Cells(row, 36).Value = ShTicket.Cells(i, 36).Value
            ' SAP anlegen
            ShZSET.Cells(row, 43).Value = "nein"
            ' CRM anlegen
            ShZSET.Cells(row, 44).Value = "ja"
            row = row + 1
        Else ' if customer(i) <> customer(i+1) then fill in template and save zset before jumping to next customer
            ' Customer
            ShZSET.Cells(6, 4).Value = ShTicket.Cells(i, 1).Value
            ' Customer description
            ShZSET.Cells(7, 4).Value = ShTicket.Cells(i, 2).Value
            ' Customer
            ShZSET.Cells(row, 3).Value = ShTicket.Cells(i, 1).Value
            ' EQ description
            ShZSET.Cells(row, 4).Value = ShTicket.Cells(i, 35).Value
            ' Material number
            ShZSET.Cells(row, 5).Value = ShTicket.Cells(i, 33).Value
            ' Serial number
            ShZSET.Cells(row, 6).Value = ShTicket.Cells(i, 34).Value
            ' Warranty start
            Year = Right(ShTicket.Cells(i, 4).Value, 4)
            Month = Mid(ShTicket.Cells(i, 4).Value, 4, 2)
            Day = Left(ShTicket.Cells(i, 4).Value, 2)
            ShZSET.Cells(row, 30).Value = Format(DateSerial(Year, Month, Day), "dd.mm.yyyy")
            ' Warranty end
            ShZSET.Cells(row, 31).Value = Format(DateSerial(Year + 1, Month, Day), "dd.mm.yyyy")
            ' Standort
            ShZSET.Cells(row, 32).Value = ShTicket.Cells(i, 3).Value
            ' Street
            ShZSET.Cells(row, 33).Value = ShTicket.Cells(i, 10).Value
            ' ZIP
            ShZSET.Cells(row, 34).Value = ShTicket.Cells(i, 14).Value
            ' City
            ShZSET.Cells(row, 35).Value = ShTicket.Cells(i, 11).Value
            ' TAG
            ShZSET.Cells(row, 36).Value = ShTicket.Cells(i, 36).Value
            ' SAP anlegen
            ShZSET.Cells(row, 43).Value = "nein"
            ' CRM anlegen
            ShZSET.Cells(row, 44).Value = "ja"
            ' save zset to another workbook
            ShZSET.Visible = xlSheetVisible
            ShZSET.Copy
            ShZSET.Visible = xlSheetVeryHidden
            Set NewBook = ActiveWorkbook
            ' save in same directory
            SavePath = ThisWorkbook.Path & "\ZSET_" & ShTicket.Cells(i, 1).Value & ".txt"
            NewBook.SaveAs SavePath, xlUnicodeText
            NewBook.Close False
            ' reset the row in template for next customer
            row = 17
            ' clear template for next customer
            ShZSET.Range("C17:AR100000").ClearContents
        End If
    Next i
    
    ShSource.Activate
    ShSource.Range("A1").Select
    MsgBox "ZSET files per customer saved in same directory of this macro", vbInformation
    
    Call ExitPoint

End Sub
Sub zzservice()
    
    Call EntryPoint
    
    ticketLR = ShTicket.Range("A" & Rows.Count).End(xlUp).row
    
    ShZZservice.Range("A2:BV100000").ClearContents
    
    ' check if country is not germany
    If ShSource.Range("A2").Value = "DE" Then
        MsgBox "ZZSERVICE will only work for countries which are not Germany!", vbCritical, "Error!"
        Call EntryPoint
        Exit Sub
    End If
    
    ' check if its not empty
    If ShSource.Range("A2").Value = "" Then
        MsgBox "No existing data to process!", vbCritical, "Error!"
        Call EntryPoint
        Exit Sub
    End If
    
    ' Customer
    ShTicket.Range("A2:A" & ticketLR).Copy
    ShZZservice.Range("A2").PasteSpecial xlPasteAll
    ' Name 1
    ShTicket.Range("C2:C" & ticketLR).Copy
    ShZZservice.Range("B2").PasteSpecial xlPasteAll
    ' Name 2
    ShTicket.Range("I2:I" & ticketLR).Copy
    ShZZservice.Range("C2").PasteSpecial xlPasteAll
    ' Street
    ShTicket.Range("J2:J" & ticketLR).Copy
    ShZZservice.Range("D2").PasteSpecial xlPasteAll
    ' City
    ShTicket.Range("K2:K" & ticketLR).Copy
    ShZZservice.Range("J2").PasteSpecial xlPasteAll
    ' District
    ShTicket.Range("L2:L" & ticketLR).Copy
    ShZZservice.Range("K2").PasteSpecial xlPasteAll
    ' Search term
    ShZZservice.Range("L2:L" & ticketLR) = "Customer"
    ' Postal Code
    ShTicket.Range("N2:N" & ticketLR).Copy
    ShZZservice.Range("M2").PasteSpecial xlPasteAll
    ' language
    ShTicket.Range("AB2:AB" & ticketLR).Copy
    ShZZservice.Range("Q2").PasteSpecial xlPasteAll
    ' country code
    ShTicket.Range("AC2:AC" & ticketLR).Copy
    ShZZservice.Range("R2").PasteSpecial xlPasteAll
    ' tax
    ShTicket.Range("AD2:AD" & ticketLR).Copy
    ShZZservice.Range("S2").PasteSpecial xlPasteAll
    ' material number
    ShTicket.Range("AG2:AG" & ticketLR).Copy
    ShZZservice.Range("Y2").PasteSpecial xlPasteAll
    ' serial number
    ShTicket.Range("AH2:AH" & ticketLR).Copy
    ShZZservice.Range("Z2").PasteSpecial xlPasteAll
    ' EQ description
    ShTicket.Range("AI2:AI" & ticketLR).Copy
    ShZZservice.Range("AA2").PasteSpecial xlPasteAll
    ' TAG
    ShTicket.Range("AJ2:AJ" & ticketLR).Copy
    ShZZservice.Range("AB2").PasteSpecial xlPasteAll
    ' description FL
    ShTicket.Range("AV2:AV" & ticketLR).Copy
    ShZZservice.Range("BC2").PasteSpecial xlPasteAll
    
    ShZZservice.Visible = xlSheetVisible
    ShZZservice.Copy
    ShZZservice.Visible = xlSheetVeryHidden
    Set NewBook = ActiveWorkbook
    SavePath = ThisWorkbook.Path & "\ZZSERVICE_" & ShSource.Range("A2").Value & ".txt"
    NewBook.SaveAs SavePath, xlUnicodeText
    NewBook.Close False
    ShSource.Activate
    ShSource.Range("A1").Select
    MsgBox "ZZSERVICE save in same directory of this macro", vbInformation
    MsgBox "Don't forget to extend consignees to both distribution channels!", vbExclamation
    
    
    Call ExitPoint
    
    
End Sub
Sub zuplxls()
    
    Dim FileToOpen As String
    Dim SelectedFile As Workbook
    Dim UserInput As String
    Dim WB As Workbook
    Dim j As Byte, r As Byte
    
    Call EntryPoint
    
    Set WB = ActiveWorkbook
    
    ShHeader.Range("A2:AH100000").ClearContents
    ShItem.Range("A2:X100000").ClearContents
    ' check if country is not Germany
    If ShSource.Range("A2") = "DE" Then
        MsgBox "This ZUPLXLS only works for countries in PA3 & PE5", vbCritical, "Error!"
        Exit Sub
    End If
    
    ' before process user needs to import zzservice after SAP updated of consignee and eq
    UserInput = MsgBox("Before proceeding with ZUPLXS please select ZZSERVICE file for EQ and Consignee searching", vbInformation + vbYesNo, "Confirmation needed")
    If UserInput = vbNo Then
        MsgBox "Process was stopped!", vbCritical
        Call ExitPoint
        Exit Sub
    Else
        FileToOpen = Application.GetOpenFilename(Filefilter:="Text files (*.txt),*txt", Title:="Select Workbook to import", MultiSelect:=False)
        'On Error GoTo here:
        Set SelectedFile = Workbooks.Open(FileToOpen)
'here:        If Err.Number = 1004 Then
            'MsgBox "Process was stopped!", vbCritical
            'Call ExitPoint
            'Exit Sub
'            End If
    End If
    
    ticketLR = ShTicket.Range("A" & Rows.Count).End(xlUp).row
    
    ' check if consignee and equipment are populated in zzservice
    If SelectedFile.Sheets(1).Range("BF2") = "" Then
        MsgBox "Looks like your zzservice file doesnt have consignees!", vbCritical
        Call ExitPoint
        SelectedFile.Close False
        Exit Sub
    Else
        If SelectedFile.Sheets(1).Range("BV2") = "" Then
            MsgBox "Looks like your zzservice file doesnt have EQ numbers!", vbCritical
            Call ExitPoint
            SelectedFile.Close False
            Exit Sub
        End If
    End If
    
    ' copy from zzservice consignees and EQ
    SelectedFile.Sheets(1).Range("BV2:BV" & ticketLR).Copy
    ShTicket.Range("BB2").PasteSpecial xlPasteAll
    SelectedFile.Sheets(1).Range("BF2:BF" & ticketLR).Copy
    ShTicket.Range("BA2").PasteSpecial xlPasteAll
    SelectedFile.Close False
    
    ShHeader.Activate
    ' loop for contract template (1 to 2 because we have HW and SW)
    sourceLR = ShSource.Range("A" & Rows.Count).End(xlUp).row
    ' (1 = HW; 2 = SW)
    For i = 1 To 2
        ' header
        For j = 2 To sourceLR
            ' statement to dont create SW for US
            If ShSource.Cells(j, 1).Value = "US" Then
                If i = 2 Then
                    MsgBox "There is no SW contracts for US", vbInformation
                    ShSource.Activate
                    Call ExitPoint
                    Exit Sub
                End If
            End If
            ' customer
            ShHeader.Cells(j, 1).Value = ShSource.Cells(j, 4).Value
            ' description
            ShHeader.Cells(j, 7).Value = ShSource.Cells(j, 6).Value
            ' consgignee
            ShHeader.Cells(j, 2).Value = Application.WorksheetFunction.XLookup(ShHeader.Cells(j, 7).Value, ShTicket.Range("C:C"), ShTicket.Range("BA:BA"))
            'division
            If i = 1 Then
               ShHeader.Cells(j, 3).Value = "SR"
            Else
                ShHeader.Cells(j, 3).Value = "R2"
            End If
            ' sales office
            ShHeader.Cells(j, 4).Value = Application.WorksheetFunction.XLookup(ShHeader.Cells(j, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("L:L"))
            ' sales group
            Select Case ShSource.Range("A" & j).Value
                Case "AU"
                    ShHeader.Cells(j, 5).Value = "AU1"
                Case "US"
                    ShHeader.Cells(j, 5).Value = "US2"
                Case "CH"
                    If i = 1 Then
                        ShHeader.Cells(j, 5).Value = "CHB"
                    Else
                        ShHeader.Cells(j, 5).Value = "CHL"
                    End If
            End Select
            'contract start
            ShHeader.Cells(j, 6).Value = Format(ShSource.Cells(j, 7).Value, "dd.mm.yyyy")
            ' billing start
            ShHeader.Cells(j, 10).Value = ShHeader.Cells(j, 6).Value
            ' group contracts
            If i = 1 Then
                ShHeader.Cells(j, 14).Value = Application.WorksheetFunction.XLookup(ShHeader.Cells(j, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("G:G"))
            Else
                ShHeader.Cells(j, 14).Value = Application.WorksheetFunction.XLookup(ShHeader.Cells(j, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("H:H"))
            End If
            ' sales rep
            ShHeader.Cells(j, 15).Value = Application.WorksheetFunction.XLookup(ShHeader.Cells(j, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("N:N"))
            ' log.emp
            ShHeader.Cells(j, 16).Value = Application.WorksheetFunction.XLookup(ShHeader.Cells(j, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("O:O"))
            ' contract end date
            ShHeader.Cells(j, 21).Value = Application.WorksheetFunction.XLookup(ShHeader.Cells(j, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("R:R"))
            'order reason
            ShHeader.Cells(j, 32).Value = "PRS"
            'PO field for CH
            If ShSource.Cells(j, 1).Value = "CH" Then
                ShHeader.Cells(j, 33).Value = Right(ShSource.Cells(j, 2).Value, 3)
            End If
        Next j
        If i = 1 Then
            SavePath = ThisWorkbook.Path & "\ZUPLXLS_Header_HW_" & ShSource.Range("A2").Value & ".txt"
        Else
            SavePath = ThisWorkbook.Path & "\ZUPLXLS_Header_SW_" & ShSource.Range("A2").Value & ".txt"
        End If
        ShHeader.Visible = xlSheetVisible
        ShHeader.Copy
        ShHeader.Visible = xlSheetVeryHidden
        Set NewBook = ActiveWorkbook
        NewBook.SaveAs SavePath, xlUnicodeText
        NewBook.Close False
        
        ' item starts here
        For r = 2 To ticketLR
            ' consignee
            ShItem.Cells(r, 1).Value = ShTicket.Cells(r, 53).Value
            ' material
            Select Case ShTicket.Cells(r, 52).Value
                Case "Card"
                If i = 1 Then
                    ShItem.Cells(r, 2).Value = "material 1"
                Else
                    ShItem.Cells(r, 2).Value = "material 2"
                End If
                Case "iCash 40"
                 If i = 1 Then
                    ShItem.Cells(r, 2).Value = "material 3"
                Else
                    ShItem.Cells(r, 2).Value = "material 4"
                End If
                Case "iCash 60"
                 If i = 1 Then
                    ShItem.Cells(r, 2).Value = "material 5"
                Else
                    ShItem.Cells(r, 2).Value = "material 6"
                End If
                Case "Attendant"
                 If i = 1 Then
                    ShItem.Cells(r, 2).Value = "material 7"
                Else
                    ShItem.Cells(r, 2).Value = "material 8"
                End If
                ' only US HW
                Case "Scanner"
                 If i = 1 Then
                    ShItem.Cells(r, 2).Value = "material 9"
                Else
                    ShItem.Cells(r, 2).Value = ""
                End If
            End Select
            ' EQ
            ShItem.Cells(r, 5).Value = ShTicket.Cells(r, 54).Value
            'SLAs
            If i = 1 Then
                ShItem.Cells(r, 7).Value = ""
                ShItem.Cells(r, 8).Value = "ALD"
                ShItem.Cells(r, 9).Value = "SC"
            Else
                ShItem.Cells(r, 7).Value = "XX"
                ShItem.Cells(r, 8).Value = ""
                ShItem.Cells(r, 9).Value = ""
            End If
        Next r
        If i = 1 Then
            SavePath = ThisWorkbook.Path & "\ZUPLXLS_Item_HW_" & ShSource.Range("A2").Value & ".txt"
        Else
            SavePath = ThisWorkbook.Path & "\ZUPLXLS_Item_SW_" & ShSource.Range("A2").Value & ".txt"
        End If
        ShItem.Visible = xlSheetVisible
        ShItem.Copy
        ShItem.Visible = xlSheetVeryHidden
        Set NewBook = ActiveWorkbook
        NewBook.SaveAs SavePath, xlUnicodeText
        NewBook.Close False
    Next i
    
    
    
    ShSource.Activate
    ShSource.Range("A1").Select
    MsgBox "ZUPLXLS files saved in same directory of this macro", vbInformation
    
    Call ExitPoint

End Sub

Sub consignee_de()
    
    Call EntryPoint
    
    sourceLR = ShSource.Range("A" & Rows.Count).End(xlUp).row
    
    ' Name 1
    ShSource.Range("F2:F" & sourceLR).Copy
    Shzgb100.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
    
    ' Street
    ShSource.Range("H2:H" & sourceLR).Copy
    Shzgb100.Range("C2").PasteSpecial xlPasteValuesAndNumberFormats
    
    ' City
    ShSource.Range("I2:I" & sourceLR).Copy
    Shzgb100.Range("D2").PasteSpecial xlPasteValuesAndNumberFormats
    
    ' Search term
    Shzgb100.Range("F2:F" & sourceLR) = "Customer"
    
    ' Postal code
    ShSource.Range("J2:J" & sourceLR).Copy
    Shzgb100.Range("G2").PasteSpecial xlPasteValuesAndNumberFormats
    
    ' language
    Shzgb100.Range("K2:K" & sourceLR) = "DE"
    
    ' country
    Shzgb100.Range("L2:L" & sourceLR) = "DE"
    
    ' tax
    Shzgb100.Range("M2:M" & sourceLR) = "B"
    
    ' save file
    SavePath = ThisWorkbook.Path & "\ZGB100_DE" & ".txt"
    Shzgb100.Visible = xlSheetVisible
    Shzgb100.Copy
    Shzgb100.Visible = xlSheetVeryHidden
    Set NewBook = ActiveWorkbook
    NewBook.SaveAs SavePath, xlUnicodeText
    NewBook.Close False
    ShSource.Activate
    MsgBox "ZGB100 saved in same directory of this workbook", vbInformation
    
    Call ExitPoint
    

End Sub


Sub new_zuplxls_de()

    Dim ShTicket As Worksheet, ShCoEq As Worksheet, ShHeader As Worksheet, ShItem As Worksheet, ShOrg As Worksheet
    Dim NewBook As Workbook
    Dim Answer As String
    Dim i As Long, j As Long, x As Long
    Dim lastRow As Long, lastRow2 As Long
    Dim FilePath As String
    
    'Set ShSource = ThisWorkbook.Worksheets("Source")
    Set ShCoEq = ThisWorkbook.Worksheets("DE_CO_EQ")
    Set ShTicket = ThisWorkbook.Worksheets("Ticket")
    Set ShHeader = ThisWorkbook.Worksheets("hh")
    Set ShItem = ThisWorkbook.Worksheets("ii")
    Set ShOrg = ThisWorkbook.Worksheets("OrgData")
    
    Call EntryPoint
    
    FilePath = ThisWorkbook.Path & "\"
    
'    ShCoEq.Visible = xlSheetVisible
    
    lastRow = ShTicket.Range("A" & Rows.Count).End(xlUp).row
    
    ShHeader.Visible = xlSheetVisible
    ShItem.Visible = xlSheetVisible
    
    ShHeader.Range("A2:AH10000").ClearContents
    ShItem.Range("A2:X10000").ClearContents
    
    
    
    ' header
    
    lastRow = ShSource.Range("A" & Rows.Count).End(xlUp).row
    lastRow2 = ShTicket.Range("A" & Rows.Count).End(xlUp).row
    
    For j = 1 To 2 ' (1 = HW ; 2 = SW)
        If j = 1 Then ' (1 = HW)
            For i = 2 To lastRow
                ' HW
                ' HH
                ShHeader.Cells(i, 1).Value = ShSource.Cells(i, 4).Value
                ShHeader.Cells(i, 2).Value = Application.WorksheetFunction.VLookup(ShSource.Cells(i, 6).Value, ShCoEq.Range("A:B"), 2, False)
                ShHeader.Cells(i, 3).Value = "SR"
                ShHeader.Cells(i, 4).Value = "D130"
                ShHeader.Cells(i, 5).Value = "4CP"
                ShHeader.Cells(i, 6).Value = Format(ShSource.Cells(i, 7).Value, "dd.mm.yyyy")
                ShHeader.Cells(i, 7).Value = ShSource.Cells(i, 6).Value
                ShHeader.Cells(i, 10).Value = ShHeader.Cells(i, 6).Value
                ShHeader.Cells(i, 14).Value = Application.WorksheetFunction.VLookup(ShHeader.Cells(i, 1).Value, ShOrg.Range("A:G"), 7, False)
                ShHeader.Cells(i, 15).Value = "586947"
                ShHeader.Cells(i, 21).Value = Application.WorksheetFunction.VLookup(ShHeader.Cells(i, 14).Value, ShOrg.Range("G:R"), 12, False)
            Next i
                ShHeader.Copy
                Set NewBook = ActiveWorkbook
                NewBook.SaveAs FilePath & "DE_Header_HW.txt", xlUnicodeText
                NewBook.Close False
                ShHeader.Range("A2:AH10000").ClearContents
            For x = 2 To lastRow2
                ' HW
                ' II
                ShItem.Cells(x, 1).Value = ShCoEq.Cells(x, 2).Value
                Select Case ShTicket.Cells(x, 54)
                    Case "Card"
                        ShItem.Cells(x, 2).Value = "material 1"
                    Case "iCash 40"
                        ShItem.Cells(x, 2).Value = "material 2"
                    Case "iCash 60"
                        ShItem.Cells(x, 2).Value = "material 3"
                    Case "Attendant"
                        ShItem.Cells(x, 2).Value = "material 4"
                End Select
                ShItem.Cells(x, 5).Value = ShCoEq.Cells(x, 4).Value
                ShItem.Cells(x, 7).Value = "ALD"
                ShItem.Cells(x, 8).Value = "ST"
                ShItem.Cells(x, 9).Value = "SC"
                ShItem.Cells(x, 19).Value = x
                Select Case ShItem.Cells(x, 2).Value
                    Case "1770062511"
                        ShItem.Cells(x, 13).Value = "100"
                    Case "1770062512"
                        ShItem.Cells(x, 13).Value = "N/A"
                    Case "1770062513"
                        ShItem.Cells(x, 13).Value = "50"
                    Case "1770062571"
                        ShItem.Cells(x, 13).Value = "10"
                End Select
            Next x
                ShItem.Copy
                Set NewBook = ActiveWorkbook
                NewBook.SaveAs FilePath & "DE_Item_HW.txt", xlUnicodeText
                NewBook.Close False
                ShItem.Range("A2:X10000").ClearContents
        Else ' (2 = SW)
            For i = 2 To lastRow
                ' HW
                ShHeader.Cells(i, 1).Value = ShSource.Cells(i, 4).Value
                ShHeader.Cells(i, 2).Value = Application.WorksheetFunction.VLookup(ShSource.Cells(i, 6).Value, ShCoEq.Range("A:B"), 2, False)
                ShHeader.Cells(i, 3).Value = "R2"
                ShHeader.Cells(i, 4).Value = "D130"
                ShHeader.Cells(i, 5).Value = "4CP"
                ShHeader.Cells(i, 6).Value = Format(ShSource.Cells(i, 7).Value, "dd.mm.yyyy")
                ShHeader.Cells(i, 7).Value = ShSource.Cells(i, 6).Value
                ShHeader.Cells(i, 10).Value = ShHeader.Cells(i, 6).Value
                ShHeader.Cells(i, 14).Value = Application.WorksheetFunction.VLookup(ShHeader.Cells(i, 1).Value, ShOrg.Range("A:H"), 8, False)
                ShHeader.Cells(i, 15).Value = "586947"
                ShHeader.Cells(i, 21).Value = Application.WorksheetFunction.VLookup(ShHeader.Cells(i, 14).Value, ShOrg.Range("H:R"), 11, False)
            Next i
                ShHeader.Copy
                Set NewBook = ActiveWorkbook
                NewBook.SaveAs FilePath & "DE_Header_SW.txt", xlUnicodeText
                NewBook.Close False
                ShHeader.Range("A2:AH10000").ClearContents
                        For x = 2 To lastRow2
                ' HW
                ' II
                ShItem.Cells(x, 1).Value = ShCoEq.Cells(x, 2).Value
                Select Case ShTicket.Cells(x, 54)
                    Case "Card"
                        ShItem.Cells(x, 2).Value = "material 1"
                    Case "iCash 40"
                        ShItem.Cells(x, 2).Value = "material 2"
                    Case "iCash 60"
                        ShItem.Cells(x, 2).Value = "material 3"
                    Case "Attendant"
                        ShItem.Cells(x, 2).Value = "material 4"
                End Select
                ShItem.Cells(x, 5).Value = ShCoEq.Cells(x, 4).Value
                ShItem.Cells(x, 7).Value = "XX"
'                ShItem.Cells(x, 8).Value = "ST"
'                ShItem.Cells(x, 9).Value = "SC"
                ShItem.Cells(x, 19).Value = x
                Select Case ShItem.Cells(x, 2).Value
                    Case "1770062522"
                        ShItem.Cells(x, 13).Value = "5"
                    Case "1770062523"
                        ShItem.Cells(x, 13).Value = "7"
                    Case "1770062561"
                        ShItem.Cells(x, 13).Value = "6"
                End Select
            Next x
                ShItem.Copy
                Set NewBook = ActiveWorkbook
                NewBook.SaveAs FilePath & "DE_Item_SW.txt", xlUnicodeText
                NewBook.Close False
                ShItem.Range("A2:X10000").ClearContents
        End If
    Next j
    
    ShCoEq.Range("A2:D100000").Clear
    
    ShHeader.Visible = xlSheetVeryHidden
    ShItem.Visible = xlSheetVeryHidden
    ShCoEq.Visible = xlSheetVeryHidden
    
    MsgBox "ZUPLXS files saved in same directory as this macro", vbInformation
    ShSource.Activate
    ShSource.Range("A1").Select
    
    
    Call ExitPoint

End Sub
