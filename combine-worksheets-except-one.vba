' Did it for a question in Stack Overflow
' Combine 4 worksheets into a new "combined" worksheet, but skip worksheet 1
' All sheets have the same table headers i.e. all sheet row 1 are the same

            Sub Combine()
              Dim J As Integer
              On Error Resume Next
              Sheets(1).Select
              Worksheets.Add
              Sheets(1).Name = "Combined"
              Sheets(3).Activate
              Range("A1").EntireRow.Select
              Selection.Copy Destination:=Sheets(1).Range("A1")
              For J = 3 To Sheets.Count
                Sheets(J).Activate
                Range("A1").Select
                Selection.CurrentRegion.Select
                Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
                Selection.Copy Destination:=Sheets(1).Range("A65536").End(xlUp)(2)
              Next
            End Sub


            Sub Combine()
              Dim Lastrow As Integer
              Dim J As Integer
              On Error Resume Next
              Sheets(1).Select
              Worksheets.Add
              Sheets(1).Name = "Combined"
              Sheets(3).Activate
              Range("A1").EntireRow.Select
              Selection.Copy Destination:=Sheets(1).Range("A1")
              For J = 3 To Sheets.Count
                Sheets(J).Activate
                ' First delete the empty rows
                Lastrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
                Range("A2:L" & Lastrow).Select
                Selection.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
                ' Then select the region as a table
                Range("A1").Select
                Selection.CurrentRegion.Select
                Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
                Selection.Copy Destination:=Sheets(1).Range("A65536").End(xlUp)(2)
              Next
            End Sub

