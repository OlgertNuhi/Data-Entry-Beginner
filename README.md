//Macros Code//
//To see the Macros Code, you need to download the file and go to Section DEVELOPER and touch MACROS//
//Sheet1 (HOME)//

Sub Reset()

Dim iRow As Long

WorksheetFunction.CountA (Sheets("Database").Range("A:A")) ' identifying the last row



With frmForm
.txtID.Value = ""
.txtName.Value = ""
.optmale.Value = False

.optFemale.Value = False


.cmbDepartment.Clear

.cmbDepartment.AddItem "HR"
.cmbDepartment.AddItem "Operation"
.cmbDepartment.AddItem "Training"
.cmbDepartment.AddItem "Quality"


.txtCity.Value = ""
.txtCountry.Value = ""


.lstDatabase.ColumnCount = 9
.lstDatabase.ColumnHeads = True

.lstDatabase.ColumnWidths = "30,60,75,40,60,45,55,70,70"

If iRow > 1 Then

.lstDatabase.RowSource = "Database!A2:I" & iRow
Else
.lstDatabase.RowSource = "Database!A2:I2"
End If

End With

End Sub

//Sheet2(DATABASE)//

Sub Submit()

Dim sh As Worksheet

Dim iRow As Long

Set sh = ThisWorkbook.Sheets("Database")

iRow = [Counta(Database!A:A)] + 1

With sh

.Cells(iRow, 1) = iRow - 1
.Cells(iRow, 2) = frmForm.txtID.Value
.Cells(iRow, 3) = frmForm.txtName.Value
.Cells(iRow, 4) = IIf(frmForm.optFemale.Value = True, "Female", "Male")
.Cells(iRow, 5) = frmForm.cmbDepartment.Value
.Cells(iRow, 6) = frmForm.txtCity.Value
.Cells(iRow, 7) = frmForm.txtCountry.Value
.Cells(iRow, 8) = Application.UserName
.Cells(iRow, 9) = [Text(Now(), "DD-MM-YYYY HH:MM:SS")]


End With

End Sub
// //
Sub Show_Form()

frmForm.Show


End Sub


