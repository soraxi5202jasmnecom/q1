Private Sub Workbook_Open()
    
    't = Application.ExecuteExcel4Macro("XL.ComputeMD5Hash(""test"")")

    Worksheets("LDR").Select
    inp = InputBox("Для открытия файла, введи пароль:", "XLTools Auth Module")
    md5 = Application.ExecuteExcel4Macro("XL.ComputeMD5Hash(""" & inp & """)")
    src = "098F6BCD4621D373CADE4E832627B4F6"
    If md5 <> src Then
        ThisWorkbook.Close
    End If
End Sub