Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

Application.DisplayAlerts = False
Application.ScreenUpdating = False

//.csv file that gets uploaded with DataSync
Workbooks.Open Filename:="F:\data_for_web.csv"

//Original Excel file that has the data
Workbooks("Data_Master.xlsm").Activate

//Activate proper tab
Worksheets("Property Data").Activate

Range("A3:BP1200").Select
Selection.Copy

Workbooks("data_for_web.csv").Activate

Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

Application.Workbooks("data_for_web.csv").Close SaveChanges:=True

Application.CutCopyMode = False
Cells(1, 1).Select

End Sub

