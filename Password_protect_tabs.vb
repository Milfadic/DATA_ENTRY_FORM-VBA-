    Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    mysheet = "BALANCES"
    If ActiveSheet.Name = mysheet Then
    Application.EnableEvents = False
    ActiveSheet.Visible = False
    response = InputBox("Enter password to view sheet")
    If response = "ENTERPASSWORD" Then
    Sheets(mysheet).Visible = True
    End If
    End If
    Application.EnableEvents = True

    mysheetB = "CODES"
    If ActiveSheet.Name = mysheetB Then
    Application.EnableEvents = False
    ActiveSheet.Visible = False
    response = InputBox("Enter password to view sheet")
    If response = "ENTERPASSWORD" Then
    Sheets(mysheetB).Visible = True
    End If
    End If
    Application.EnableEvents = True

    End Sub





