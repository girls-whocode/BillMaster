# BillMaster

An advanced Bill organizer with Pay down system. Although this is still a work in progress, almost all features are working. There are a few bugs that I will be working out over time. If you would like to see the current bugs, please look at the issues section and use the bugs filter.

## Prerequisites
First and foremost, this script uses VBA, on all versions of Excel, when VBA or Macros are used, a warning will appear at the top of the screen. This template uses VBA for the creation of a new account by copying a template, then changing the Data Index to add the new Bill. The VBA code is:

```vb
Private Sub CommandButton1_Click()
    Dim xName As String
    Dim xSht As Object
    Dim xNWS As Worksheet

    ' On Error Resume Next
        xType = Range("B4").Value ' Get the selection of Credit, Bill, Investment, or Cash
        If xType = "Credit Accounts" Then
            xName = Application.InputBox("Credit Name ", "Credit Account")
            If xName = "" Then
                MsgBox ("User Canceled or Nothing")
                Exit Sub
            End If
        ElseIf xType = "Bill Accounts" Then
            xName = Application.InputBox("Bill Name ", "Bill Account", "")
            If xName = "" Then
                MsgBox ("User Canceled or Nothing")
                Exit Sub
            End If
        ElseIf xType = "Investments" Then
            xName = Application.InputBox("Investment Name ", "Investment Account", "")
            If xName = "" Then
                MsgBox ("User Canceled or Nothing")
                Exit Sub
            End If
        ElseIf xType = "Cash Accounts" Then
            xName = Application.InputBox("Cash Name ", "Cash Account", "")
            If xName = "" Then
                MsgBox ("User Canceled or Nothing")
                Exit Sub
            End If
        End If
        
        ' Depending on the selection, activate the sheet, copy the sheet, make sure it is visable
        ' Give it the name from the message box, change the Data's Index value from 9 to 8 and back
        ' to 9 so it will recalculate the page automatically.
        If xType = "Credit Accounts" Then
            Worksheets("CredTemplate").Activate
            ActiveSheet.Copy after:=Sheets(Sheets.Count)
            Set xNWS = Sheets(Sheets.Count)
            xNWS.Visible = True
            xNWS.Name = xName
            Sheets("Data").Cells(3, 11).Value = 8
            Sheets("Data").Cells(3, 11).Value = 9
            Worksheets(xName).Activate
        ElseIf xType = "Bill Accounts" Then
            Worksheets("BillTemplate").Activate
            ActiveSheet.Copy after:=Sheets(Sheets.Count)
            Set xNWS = Sheets(Sheets.Count)
            xNWS.Visible = True
            xNWS.Name = xName
            Sheets("Data").Cells(3, 11).Value = 8
            Sheets("Data").Cells(3, 11).Value = 9
            Worksheets(xName).Activate
        ElseIf xType = "Investments" Then
            Worksheets("InvestTemplate").Activate
            ActiveSheet.Copy after:=Sheets(Sheets.Count)
            Set xNWS = Sheets(Sheets.Count)
            xNWS.Visible = True
            xNWS.Name = xName
            Sheets("Data").Cells(3, 11).Value = 8
            Sheets("Data").Cells(3, 11).Value = 9
            Worksheets(xName).Activate
        ElseIf xType = "Cash Accounts" Then
            Worksheets("CashTemplate").Activate
            ActiveSheet.Copy after:=Sheets(Sheets.Count)
            Set xNWS = Sheets(Sheets.Count)
            xNWS.Visible = True
            xNWS.Name = xName
            Sheets("Data").Cells(3, 11).Value = 8
            Sheets("Data").Cells(3, 11).Value = 9
            Worksheets(xName).Activate
        End If
End Sub
```

## To start

BillMaster has been designed for people without Excel or computer experience to be able to create, modify, and use almost all features of this system. Open the file with Microsoft Excel Office 365 or newer. 

NOTE: This system uses dynamic tables that was added in Office 365.
