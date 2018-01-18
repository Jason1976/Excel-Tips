Sub Combine()
    Dim J As Integer
    
    On Error Resume Next
    Sheets(1).Select
    Worksheets.Add
    Sheets(1).Name = "Combined"
    Sheets(2).Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets(1).Range("A1")
    
    For J = 2 To Sheets.Count
        Sheets(J).Activate
        Range("A1").Select
        Selection.CurrentRegion.Select
        Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
        Selection.Copy Destination:=Sheets(1).Range("A65536").End(xlUp)(2)
        //只需要複製欄位值，請註解掉上一行程式碼，並將下列二行的註解拿掉
        //Selection.Copy
        //Sheets(1).Range("A65536").End(xlUp)(2).PasteSpecial xlPasteValues
    Next
End Sub
