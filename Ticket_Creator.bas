Attribute VB_Name = "Ticket_Creator"
Sub Generate_Data_for_ticket()

    ' After requester fills in Store Info Table macro generates New Store Form
    
    Dim tbLastRow As Byte
    Dim i As Byte, j As Byte
    Dim nSCO As Byte
    Dim row As Long
    Dim SCOindex As Byte
    Dim MaterialType As String
    
    Call EntryPoint
    
    ' Check if there is at least 1 store in table
    If ShSource.Cells(2, 1).Value = NullString Then
        MsgBox "Please populate at least 1 store in the Store Info Table!", vbCritical, "Error!"
    End If
    
    ' Clear form before creating new one
    ShTicket.Range("A2:BB100000").ClearContents
    
    ' Check Store Info Table last filled row
    tbLastRow = ShSource.Range("A" & Rows.Count).End(xlUp).row
    
    ' Need double loop for each row(store) and devices per row (store)
    row = 2
    For i = 2 To tbLastRow
        nSCO = ShSource.Cells(i, 12).Value
        ' Store might have already installed previously other device - index will tell us from which number is starts
        SCOindex = ShSource.Cells(i, 15).Value
        MaterialType = ShSource.Cells(i, 13).Value
        For j = 1 To nSCO
            ' Customer
            ShTicket.Cells(row, 1).Value = ShSource.Cells(i, 4).Value
            ' Customer Name
            ShTicket.Cells(row, 2).Value = Application.WorksheetFunction.XLookup(ShTicket.Cells(row, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("C:C"))
            ' Description for contract
            ShTicket.Cells(row, 3).Value = "Customer Name " & Right(ShSource.Cells(i, 6).Value, 8)
            ' Contract Start
            ShTicket.Cells(row, 4).Value = Format(ShSource.Cells(i, 7).Value, "dd.mm.yyyy")
            ' Invoice Start
            ShTicket.Cells(row, 5).Value = ShTicket.Cells(row, 4).Value
            ' Name 1 Consignee
            ShTicket.Cells(row, 8).Value = ShTicket.Cells(row, 3).Value
            ' Street
            ShTicket.Cells(row, 10).Value = ShSource.Cells(i, 8).Value
            ' City
            ShTicket.Cells(row, 11).Value = ShSource.Cells(i, 9).Value
            ' Search Term
            ShTicket.Cells(row, 13).Value = "Customer"
            ' Postal Code
            ShTicket.Cells(row, 14).Value = ShSource.Cells(i, 10).Value
            'Language/Country/TAX/Name 2 for AU/State for US & AU
            Select Case ShSource.Range("A" & i).Value
                Case "US"
                    ShTicket.Cells(row, 28).Value = "EN"
                    ShTicket.Cells(row, 29).Value = "US"
                    ShTicket.Cells(row, 30).Value = "B"
                    ShTicket.Cells(row, 12).Value = Application.WorksheetFunction.XLookup(ShTicket.Cells(row, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("F:F"))
                Case "AU"
                    ShTicket.Cells(row, 28).Value = "EN"
                    ShTicket.Cells(row, 29).Value = "AU"
                    ShTicket.Cells(row, 30).Value = "B"
                    ShTicket.Cells(row, 9).Value = "ALDI " & ShTicket.Cells(row, 11).Value
                    ShTicket.Cells(row, 12).Value = Application.WorksheetFunction.XLookup(ShTicket.Cells(row, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("F:F"))
                Case "CH"
                    ShTicket.Cells(row, 28).Value = "DE"
                    ShTicket.Cells(row, 29).Value = "CH"
                    ShTicket.Cells(row, 30).Value = "1"
                Case "DE"
                    ShTicket.Cells(row, 28).Value = "DE"
                    ShTicket.Cells(row, 29).Value = "DE"
                    ShTicket.Cells(row, 30).Value = "B"
            End Select
            ' Material Number
            Select Case MaterialType
                Case "Card"
                   ShTicket.Cells(row, 33).Value = "material 1"
                Case "iCash 40"
                    ShTicket.Cells(row, 33).Value = "material 2"
                Case "iCash 60"
                    ShTicket.Cells(row, 33).Value = "material 3"
            End Select
            ' Equipment Description
            ShTicket.Cells(row, 35).Value = Application.WorksheetFunction.XLookup(ShTicket.Cells(row, 33).Value, ShLists.Range("K:K"), ShLists.Range("L:L"))
            ' TAG different for Card and Cash systems
            If ShSource.Cells(i, 13).Value = "Card" Then
                ShTicket.Cells(row, 36).Value = ShSource.Range("A" & i).Value & "ALS" & Right(ShSource.Cells(i, 6).Value, 6) & "SCO0" & j + SCOindex - 1
            Else
                ShTicket.Cells(row, 36).Value = ShSource.Range("A" & i).Value & "ALS" & Right(ShSource.Cells(i, 6).Value, 6) & "SCC0" & j + SCOindex - 1
            End If
            ' To comply with allowed TAG Len in the project
            If Len(ShTicket.Cells(row, 36).Value) > 16 Then
                ShTicket.Cells(row, 36).Value = Replace(ShTicket.Cells(row, 36).Value, "SCO0", "SCO")
            End If
            ' Functional Location
            ShTicket.Cells(row, 48).Value = ShTicket.Cells(row, 3).Value
            ' material type on columns AZ
            ShTicket.Cells(row, 52).Value = ShSource.Cells(i, 13).Value
            row = row + 1
            'Additional device in case of USA
            If ShSource.Cells(2, 1).Value = "US" Then
            ' Customer
                ShTicket.Cells(row, 1).Value = ShSource.Cells(i, 4).Value
                ' Customer Name
                ShTicket.Cells(row, 2).Value = Application.WorksheetFunction.XLookup(ShTicket.Cells(row, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("C:C"))
                ' Description for contract
                ShTicket.Cells(row, 3).Value = "Customer " & Right(ShSource.Cells(i, 6).Value, 8)
                ' Contract Start
                ShTicket.Cells(row, 4).Value = Format(ShSource.Cells(i, 7).Value, "dd.mm.yyyy")
                ' Invoice Start
                ShTicket.Cells(row, 5).Value = ShTicket.Cells(row, 4).Value
                ' Name 1 Consignee
                ShTicket.Cells(row, 8).Value = ShTicket.Cells(row, 3).Value
                ' Street
                ShTicket.Cells(row, 10).Value = ShSource.Cells(i, 8).Value
                ' City
                ShTicket.Cells(row, 11).Value = ShSource.Cells(i, 9).Value
                ' Search Term
                ShTicket.Cells(row, 13).Value = "Customer"
                ' Postal Code
                ShTicket.Cells(row, 14).Value = ShSource.Cells(i, 10).Value
                'Language/Country/TAX/Name 2 for AU/State for US & AU
                Select Case ShSource.Range("A" & i).Value
                    Case "US"
                        ShTicket.Cells(row, 28).Value = "EN"
                        ShTicket.Cells(row, 29).Value = "US"
                        ShTicket.Cells(row, 30).Value = "B"
                        ShTicket.Cells(row, 12).Value = Application.WorksheetFunction.XLookup(ShTicket.Cells(row, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("F:F"))
                    Case "AU"
                        ShTicket.Cells(row, 28).Value = "EN"
                        ShTicket.Cells(row, 29).Value = "AU"
                        ShTicket.Cells(row, 30).Value = "B"
                        ShTicket.Cells(row, 9).Value = "ALDI " & ShTicket.Cells(row, 11).Value
                        ShTicket.Cells(row, 12).Value = Application.WorksheetFunction.XLookup(ShTicket.Cells(row, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("F:F"))
                    Case "CH"
                        ShTicket.Cells(row, 28).Value = "DE"
                        ShTicket.Cells(row, 29).Value = "CH"
                        ShTicket.Cells(row, 30).Value = "1"
                    Case "DE"
                        ShTicket.Cells(row, 28).Value = "DE"
                        ShTicket.Cells(row, 29).Value = "DE"
                        ShTicket.Cells(row, 30).Value = "B"
                End Select
                'Material Number
                ShTicket.Cells(row, 33).Value = "material 4"
                ' Equipment Description
                ShTicket.Cells(row, 35).Value = Application.WorksheetFunction.XLookup(ShTicket.Cells(row, 33).Value, ShLists.Range("K:K"), ShLists.Range("L:L"))
                ' TAG
                ShTicket.Cells(row, 36).Value = ShSource.Range("A" & i).Value & "ALS" & Right(ShSource.Cells(i, 6).Value, 6) & "SSL0" & j + SCOindex - 1
                ' To comply with allowed TAG Len in the project
                If Len(ShTicket.Cells(row, 36).Value) > 16 Then
                    ShTicket.Cells(row, 36).Value = Replace(ShTicket.Cells(row, 36).Value, "SSL0", "SSL")
                End If
                ' Functional Location
                ShTicket.Cells(row, 48).Value = ShTicket.Cells(row, 3).Value
                ' material type on columns AZ
                ShTicket.Cells(row, 52).Value = "Scanner"
                row = row + 1
            End If
        Next j
        ' for the additional PAS device
        If ShSource.Cells(i, 14).Value = "YES" Then
            ' Customer
            ShTicket.Cells(row, 1).Value = ShSource.Cells(i, 4).Value
            ' Customer Name
            ShTicket.Cells(row, 2).Value = Application.WorksheetFunction.XLookup(ShTicket.Cells(row, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("C:C"))
            ' Description for contract
            ShTicket.Cells(row, 3).Value = "ALDI " & Right(ShSource.Cells(i, 6).Value, 8)
            ' Contract Start
            ShTicket.Cells(row, 4).Value = Format(ShSource.Cells(i, 7).Value, "dd.mm.yyyy")
            ' Invoice Start
            ShTicket.Cells(row, 5).Value = ShTicket.Cells(row, 4).Value
            ' Name 1 Consignee
            ShTicket.Cells(row, 8).Value = ShTicket.Cells(row, 3).Value
            ' Street
            ShTicket.Cells(row, 10).Value = ShSource.Cells(i, 8).Value
            ' City
            ShTicket.Cells(row, 11).Value = ShSource.Cells(i, 9).Value
            ' Search Term
            ShTicket.Cells(row, 13).Value = "Customer"
            ' Postal Code
            ShTicket.Cells(row, 14).Value = ShSource.Cells(i, 10).Value
            'Language/Country/TAX/Name 2 for AU/State for US & AU
            Select Case ShSource.Range("A" & i).Value
                  Case "US"
                      ShTicket.Cells(row, 28).Value = "EN"
                      ShTicket.Cells(row, 29).Value = "US"
                      ShTicket.Cells(row, 30).Value = "B"
                      ShTicket.Cells(row, 12).Value = Application.WorksheetFunction.XLookup(ShTicket.Cells(row, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("F:F"))
                  Case "AU"
                      ShTicket.Cells(row, 28).Value = "EN"
                      ShTicket.Cells(row, 29).Value = "AU"
                      ShTicket.Cells(row, 30).Value = "B"
                      ShTicket.Cells(row, 9).Value = "Customer " & ShTicket.Cells(row, 11).Value
                      ShTicket.Cells(row, 12).Value = Application.WorksheetFunction.XLookup(ShTicket.Cells(row, 1).Value, ShOrg.Range("A:A"), ShOrg.Range("F:F"))
                  Case "CH"
                      ShTicket.Cells(row, 28).Value = "DE"
                      ShTicket.Cells(row, 29).Value = "CH"
                      ShTicket.Cells(row, 30).Value = "1"
                  Case "DE"
                      ShTicket.Cells(row, 28).Value = "DE"
                      ShTicket.Cells(row, 29).Value = "DE"
                      ShTicket.Cells(row, 30).Value = "B"
              End Select
            'Material Number
            ShTicket.Cells(row, 33).Value = "material 5"
            ' Equipment Description
            ShTicket.Cells(row, 35).Value = Application.WorksheetFunction.XLookup(ShTicket.Cells(row, 33).Value, ShLists.Range("K:K"), ShLists.Range("L:L"))
            ' TAG
            ShTicket.Cells(row, 36).Value = ShSource.Range("A" & i).Value & "ALS" & Right(ShSource.Cells(i, 6).Value, 6) & "PAS01"
            ' Functional Location
            ShTicket.Cells(row, 48).Value = ShTicket.Cells(row, 3).Value
            ' material type on columns AZ
            ShTicket.Cells(row, 52).Value = "Attendant"
            row = row + 1
        End If
    Next i
    
    ' Check for duplicates in serial
    ShTicket.Activate
    ShTicket.Columns("AH:AH").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    MsgBox "Template created, please fill in serial numbers in column AH", vbInformation
    ShTicket.Range("AH2").Select
    Call ExitPoint
    
End Sub
