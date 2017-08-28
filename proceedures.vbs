Option Explicit
Sub Button_import()

Dim file_list() As Variant
Dim file_count As Long
Dim file_path As String

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
sheet_antemp.Visible = xlSheetVisible

file_path = cell_filepath.Address

Call check_path(file_path)
file_list = ListFiles(file_path)
file_count = UBound(file_list) + 1

Call ImportWorksheets(file_path, file_list)
Call Print_filelist(file_list)

Application.Calculation = xlCalculationAutomatic
sheet_antemp.Visible = xlSheetHidden
Application.ScreenUpdating = True

sheet_datainput.Range("C1").Select
Application.GoTo reference:=ActiveCell, scroll:=True

'Cleanup
Erase file_list

MsgBox ("Analysis Complete!")
End Sub
Sub Button_reset()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim i As Long, arr_el As Long, ds_num As Long
Dim ds_names() As String

On Error GoTo clean_sheet

'Delete data sheets
ds_num = 0

For i = 1 To Worksheets.Count
If Left(Sheets(i).Name, 7) = "DataSet" Then
ds_num = ds_num + 1
End If
Next i

ReDim ds_names(ds_num - 1)

arr_el = 0
For i = 1 To Worksheets.Count
If Left(Sheets(i).Name, 7) = "DataSet" Then
ds_names(arr_el) = Sheets(i).Name
arr_el = arr_el + 1
End If
Next i

For i = 0 To UBound(ds_names)
Application.DisplayAlerts = False
ThisWorkbook.Sheets(ds_names(i)).Delete
Application.DisplayAlerts = True
Next i

'Delete data entries from result table

sheet_datainput.Range("E3:N" & sheet_datainput.Range("E3").End(xlDown).Row).Clear

sheet_datainput.Range("A1").Select
Application.GoTo reference:=ActiveCell, scroll:=True


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

MsgBox ("The sheet has been reset!")
Exit Sub
clean_sheet:
MsgBox ("This sheet does not need a reset!")
End Sub
Sub toggle_instructions()
If Worksheets("Instructions").Visible = xlSheetVisible Then
Worksheets("Instructions").Visible = xlSheetHidden
Worksheets("Data_input").Activate
Else
Worksheets("Instructions").Visible = xlSheetVisible
Worksheets("Instructions").Activate
End If

End Sub
Sub toggle_datatemplate()
If Worksheets("Example_Data").Visible = xlSheetVisible Then
Worksheets("Example_Data").Visible = xlSheetHidden
Worksheets("Instructions").Activate
Else
Worksheets("Example_Data").Visible = xlSheetVisible
Worksheets("Example_Data").Activate
End If

End Sub



Sub Button_Recalculate()
    Dim iCount As Integer, i As Integer
    Dim sText As String
    Dim lNum As String
    Dim sheet_num As Long
   
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
   
   'Identify sheet number
    sText = ActiveSheet.Name
    For iCount = Len(sText) To 1 Step -1
        If IsNumeric(Mid(sText, iCount, 1)) Then
            i = i + 1
            lNum = Mid(sText, iCount, 1) & lNum
        End If
         
        If i = 1 Then lNum = CInt(Mid(lNum, 1, 1))
    Next iCount
    sheet_num = CLng(lNum)
    
    clear_template (sheet_num)
    Adjust_Template (sheet_num)
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

MsgBox ("Recalculation complete!")

End Sub

Sub copy_global(sheet_num As Long)
'Copy Global Values
'A total
sheet_datainput.Range("B4").Copy destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H34")
'C total
sheet_datainput.Range("B5").Copy destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H35")
'Quantum yield ratio
sheet_datainput.Range("B6").Copy destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H36")
'Work Area start
sheet_datainput.Range("B10").Copy destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H37")
'Background region 1
sheet_datainput.Range("B11").Copy destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H38")
'Background region 2
sheet_datainput.Range("B12").Copy destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H39")
'Work Area end
sheet_datainput.Range("B13").Copy destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H40")
'Sigma
sheet_datainput.Range("B7").Copy destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H41")
'Propagation Time
sheet_datainput.Range("B8").Copy destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("G50")
'Peak width A
sheet_datainput.Range("B15").Copy destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H42")
'Peak width C
sheet_datainput.Range("B16").Copy destination:=ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H43")

End Sub
Sub clear_template(sheet_num As Long)

ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("I52:I" & sheet_datainput.Range("I52").End(xlDown).Row).Clear
ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("J50:J" & sheet_datainput.Range("J50").End(xlDown).Row).Clear
ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("L52:L" & sheet_datainput.Range("L52").End(xlDown).Row).Clear
ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("M51:M" & sheet_datainput.Range("M51").End(xlDown).Row).Clear
ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("Q51:Q" & sheet_datainput.Range("Q51").End(xlDown).Row).Clear
ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("R50:R" & sheet_datainput.Range("R50").End(xlDown).Row).Clear
ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("S50:S" & sheet_datainput.Range("S50").End(xlDown).Row).Clear

ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("V51:V" & sheet_datainput.Range("V51").End(xlDown).Row).Clear
ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("W51:W" & sheet_datainput.Range("W51").End(xlDown).Row).Clear

ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("Z50:Z" & sheet_datainput.Range("Z50").End(xlDown).Row).Clear

ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("AA51:AI" & sheet_datainput.Range("AA51").End(xlDown).Row).Clear
ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("AK50:AS" & sheet_datainput.Range("A50").End(xlDown).Row).Clear

End Sub

Sub Adjust_Template(sheet_num As Long)
Dim total_points As Long, segment_start As Long, segment_end As Long, segment_point_count As Long, total_bg_interval As Long, max_a_points As Long, max_c_points As Long
Dim sampling_rate As Double
Dim cur_sheet As Worksheet

Set cur_sheet = ThisWorkbook.Sheets("DataSet_" & sheet_num)

cur_sheet.Range("H34:J43").Calculate

total_points = cur_sheet.Range("B9").Value
sampling_rate = cur_sheet.Range("B8").Value
segment_start = cur_sheet.Range("J37").Value
segment_end = cur_sheet.Range("J40").Value

segment_point_count = (segment_end - segment_start) * sampling_rate
total_bg_interval = (cur_sheet.Range("J38").Value + cur_sheet.Range("J39").Value) * sampling_rate

'Application.Calculation = xlCalculationManual

'Initial concentrations
cur_sheet.Range("H34").Value = "=OFFSET(Data_input!G3," & sheet_num - 1 & ",0)"
cur_sheet.Range("H35").Value = "=OFFSET(Data_input!H3," & sheet_num - 1 & ",0)"

'Fill out dynamic ranges
'time, s column
cur_sheet.Range("I51").Copy destination:=cur_sheet.Range("I52:I" & total_points + 49)
'orig column
cur_sheet.Range("J50:J" & total_points + 49).FormulaArray = "=OFFSET($A$1,13,,$B$9)"
'index column
cur_sheet.Range("L51").Copy destination:=cur_sheet.Range("L52:L" & segment_point_count + 49)
'time min column
cur_sheet.Range("M50").Copy destination:=cur_sheet.Range("M51:M" & total_points + 49)
'time_window column
cur_sheet.Range("R50:R" & segment_point_count + 49).FormulaArray = "=OFFSET($I$50,($J$37+$H$50)*$Q$41,,($J$40-$J$37)*$Q$41+1)"
'time_window_min column
cur_sheet.Range("Q50").Copy destination:=cur_sheet.Range("Q51:Q" & segment_point_count + 49)
'RFU_window column
cur_sheet.Range("S50:S" & segment_point_count + 49).FormulaArray = "=OFFSET($J$50,($J$37+$H$50)*$Q$41,,($J$40-$J$37)*$Q$41+1)"
'bg_time
cur_sheet.Range("V50").Copy destination:=cur_sheet.Range("V51:V" & total_bg_interval + 49)
'bg_RFU
cur_sheet.Range("W50").Copy destination:=cur_sheet.Range("W51:W" & total_bg_interval + 49)
'RFU_corr
cur_sheet.Range("Z50:Z" & segment_point_count + 49).FormulaArray = "=RFU_window-(time_window*$U$35+$V$35)"
'RFU_rev
cur_sheet.Range("AA50").Copy destination:=cur_sheet.Range("AA51:AA" & segment_point_count + 49)
'Divergency
cur_sheet.Range("AC50").Copy destination:=cur_sheet.Range("AC51:AC" & segment_point_count + 49)
'max1
cur_sheet.Range("AD50").Copy destination:=cur_sheet.Range("AD51:AD" & segment_point_count + 49)
'tc,st
cur_sheet.Range("AE50").Copy destination:=cur_sheet.Range("AE51:AE" & segment_point_count + 49)
'Div_rev
cur_sheet.Range("AG50").Copy destination:=cur_sheet.Range("AG51:AG" & segment_point_count + 49)
'max2
cur_sheet.Range("AH50").Copy destination:=cur_sheet.Range("AH51:AH" & segment_point_count + 49)
'ta,end
cur_sheet.Range("AI50").Copy destination:=cur_sheet.Range("AI51:AI" & segment_point_count + 49)


'max1_RFU
cur_sheet.Range("AK50:AK" & (segment_point_count / 2) + 49).FormulaArray = "=OFFSET($Z$50,($AD$34-$J$37)*$Q$41-$H$42*$AD$38+1,,$H$42*$AD$38*2)"
'max1_time
cur_sheet.Range("AL50:AL" & (segment_point_count / 2) + 49).FormulaArray = "=OFFSET($R$50,($AD$34-$J$37)*$Q$41-$H$42*$AD$38+1,,$H$42*$AD$38*2)"
'max1_time2
cur_sheet.Range("AM50:AM" & (segment_point_count / 2) + 49).FormulaArray = "=max1_time^2"
'max1_fit
cur_sheet.Range("AN50:AN" & (segment_point_count / 2) + 49).FormulaArray = "=AL50:AL350^2*$AK$35+AL50:AL350*$AL$35+$AM$35"

'max2_RFU
cur_sheet.Range("AP50:AP" & (segment_point_count / 2) + 49).FormulaArray = "=OFFSET($Z$50,($AH$34-$J$37)*$Q$41-$H$43*$AH$38+1,,$H$43*$AH$38*2)"
'max2_time
cur_sheet.Range("AQ50:AQ" & (segment_point_count / 2) + 49).FormulaArray = "=OFFSET($R$50,($AH$34-$J$37)*$Q$41-$H$43*$AH$38+1,,$H$43*$AH$38*2)"
'max2_time2
cur_sheet.Range("AR50:AR" & (segment_point_count / 2) + 49).FormulaArray = "=max2_time^2"
'max2_fit
cur_sheet.Range("AS50:AS" & (segment_point_count / 2) + 49).FormulaArray = "=AQ50:AQ350^2*$AP$35+AQ50:AQ350*$AQ$35+$AR$35"


'Update charts

cur_sheet.ChartObjects("Chart 1").Activate
ActiveChart.SeriesCollection(1).XValues = cur_sheet.Range("M50:M" & total_points + 49)
ActiveChart.SeriesCollection(1).Values = cur_sheet.Range("J50:J" & total_points + 49)

cur_sheet.ChartObjects("Chart 2").Activate
ActiveChart.SeriesCollection(1).XValues = cur_sheet.Range("R50:R" & segment_point_count + 49)
ActiveChart.SeriesCollection(1).Values = cur_sheet.Range("Z50:Z" & segment_point_count + 49)
ActiveChart.Axes(xlCategory).MinimumScale = segment_start - (segment_end - segment_start) * 0.1
ActiveChart.Axes(xlCategory).MaximumScale = segment_end + (segment_end - segment_start) * 0.1


cur_sheet.ChartObjects("Chart 3").Activate
ActiveChart.SeriesCollection(1).XValues = cur_sheet.Range("Q50:Q" & segment_point_count + 49)
ActiveChart.SeriesCollection(1).Values = cur_sheet.Range("Z50:Z" & segment_point_count + 49)
ActiveChart.Axes(xlCategory).MinimumScale = (segment_start / 60)
ActiveChart.Axes(xlCategory).MaximumScale = (segment_end / 60)


cur_sheet.ChartObjects("Chart 4").Activate
ActiveChart.SeriesCollection(1).XValues = cur_sheet.Range("AL50:AL" & (segment_point_count / 2) + 49)
ActiveChart.SeriesCollection(1).Values = cur_sheet.Range("AK50:AK" & (segment_point_count / 2) + 49)
ActiveChart.SeriesCollection(2).XValues = cur_sheet.Range("AL50:AL" & (segment_point_count / 2) + 49)
ActiveChart.SeriesCollection(2).Values = cur_sheet.Range("AN50:AN" & (segment_point_count / 2) + 49)


cur_sheet.ChartObjects("Chart 5").Activate
ActiveChart.SeriesCollection(1).XValues = cur_sheet.Range("AQ50:AQ" & (segment_point_count / 2) + 49)
ActiveChart.SeriesCollection(1).Values = cur_sheet.Range("AP50:AP" & (segment_point_count / 2) + 49)
ActiveChart.SeriesCollection(2).XValues = cur_sheet.Range("AQ50:AQ" & (segment_point_count / 2) + 49)
ActiveChart.SeriesCollection(2).Values = cur_sheet.Range("AS50:AS" & (segment_point_count / 2) + 49)

cur_sheet.ChartObjects("Chart 6").Activate
ActiveChart.SeriesCollection(1).XValues = cur_sheet.Range("R50:R" & segment_point_count + 49)
ActiveChart.SeriesCollection(1).Values = cur_sheet.Range("AC50:AC" & segment_point_count + 49)
ActiveChart.Axes(xlCategory).MinimumScale = (segment_start)
ActiveChart.Axes(xlCategory).MaximumScale = (segment_end)


'Application.Calculation = xlCalculationAutomatic

End Sub

Sub check_path(file_path As String)
'This Sub corrects folder path if entered incorrectly
    Dim file_path_val As String

    file_path_val = Range(file_path).Value

    If Right(file_path_val, 1) <> "\" Then
        Range(file_path).Value = Range(file_path).Value & "\"
    End If
End Sub
Function ListFiles(file_path As String) As Variant
'This Function gets the list of files in a folder
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim file_num As Long
    Dim file_list() As Variant
        
   'Application.Calculation = xlCalculationManual
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Get the folder object associated with the directory
    Set objFolder = objFSO.GetFolder(Range(file_path).Value)

    'Loop through the Files collection
    file_num = 0
    ReDim file_list(objFolder.Files.Count - 1)
    
    For Each objFile In objFolder.Files
        file_list(file_num) = objFile.Name
        file_num = file_num + 1
    Next
     
     'Clean up!
    Set objFolder = Nothing
    Set objFile = Nothing
    Set objFSO = Nothing
    
    'Application.Calculation = xlCalculationAutomatic
    
    ListFiles = file_list
     
End Function
Sub Print_filelist(file_list As Variant)
 'This Sub prints the list of files in a folder in the result table
Dim i As Long

sheet_datainput.Select
'Application.Calculation = xlCalculationManual

For i = 0 To UBound(file_list)
'Filenumber
Cells(i + 3, "E").Value = i + 1
'Filename
Cells(i + 3, "F").Value = file_list(i)
'Initial Concentration A
sheet_datainput.Range("B4").Copy
sheet_datainput.Cells(i + 3, "G").PasteSpecial xlPasteValues
'Initial Concentration B
sheet_datainput.Range("B5").Copy
sheet_datainput.Cells(i + 3, "H").PasteSpecial xlPasteValues
'Kd
Cells(i + 3, "I").Value = "=DataSet_" & i + 1 & "!H44"
'kon
Cells(i + 3, "J").Value = "=DataSet_" & i + 1 & "!H46"
'koff
Cells(i + 3, "K").Value = "=DataSet_" & i + 1 & "!H45"
'A
Cells(i + 3, "L").Value = "=DataSet_" & i + 1 & "!N38"
'C-C*
Cells(i + 3, "M").Value = "=DataSet_" & i + 1 & "!N40"
'D+A*+C*
Cells(i + 3, "N").Value = "=DataSet_" & i + 1 & "!N42"


'Number format
sheet_datainput.Range(sheet_datainput.Cells(i + 3, "G"), sheet_datainput.Cells(i + 3, "K")).NumberFormat = "0.00E+00"
sheet_datainput.Range(sheet_datainput.Cells(i + 3, "L"), sheet_datainput.Cells(i + 3, "N")).NumberFormat = "0%"
Next i
'Application.Calculation = xlCalculationAutomatic
End Sub
Sub ImportWorksheets(file_path As String, file_list As Variant)
' This macro will import multiple files into this workbook
Dim i As Long

For i = 0 To UBound(file_list)
Call ImportSheet(file_path, file_list(i), i + 1)
Application.Wait (Now + #12:00:01 AM#)
Next i
    
End Sub
Sub ImportSheet(file_path As String, filename As Variant, sheet_num As Long)
    ' This macro will import a file into this workbook
    Dim source As Variant
    
    'Application.Calculation = xlCalculationManual
    
    sheet_antemp.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Worksheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Worksheets.Count).Name = "DataSet_" & sheet_num
    ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("H2").Value = filename
    
    On Error GoTo fail_files
    Set source = Application.Workbooks.Open(Range(file_path).Value & filename, ReadOnly:=True)
    
    On Error GoTo 0
    Application.DisplayAlerts = False
    'Copy from source sheet
     Windows(filename).Activate
     source.Sheets(1).Range("A1:E" & source.Sheets(1).Range("A14").End(xlDown).Row).Select
     Selection.Copy
     
     
     ThisWorkbook.Sheets("DataSet_" & sheet_num).Activate
     ThisWorkbook.Sheets("DataSet_" & sheet_num).Range("A1").PasteSpecial xlPasteValues
     
        
    Windows(filename).Activate
    ActiveWorkbook.Close SaveChanges:=False
    sheet_datainput.Activate
    
    Application.DisplayAlerts = True
    'Application.Calculation = xlCalculationAutomatic
    
    clear_template (sheet_num)
    copy_global (sheet_num)
    Adjust_Template (sheet_num)
    
    Exit Sub
' In case of error
fail_files:
MsgBox "Problem opening files. Please verify the file path."

End Sub

