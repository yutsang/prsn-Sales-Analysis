VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5892
   ClientLeft      =   100
   ClientTop       =   460
   ClientWidth     =   10780
   OleObjectBlob   =   "Sales Analysis - Userform 2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strDir As String, wksheet As Worksheet, wkbook As Workbook, Summary_Sheet As Worksheet, Summary_Book As Workbook, forecast_book As Workbook, forecast_sheet As Worksheet


Private Sub btnClose_Click()

Dim intlist As Integer

'check if the user selected a workbook
If wkbook Is Nothing Then

    MsgBox "Please select a workbook!"

    Else

'ask for comfirm
        If MsgBox("Are you sure you want to close this workbook?", vbYesNo, "Confirm") = vbYes Then

            ThisWorkbook.Activate
        
            lstDirWB.AddItem wkbook.Name

            Application.DisplayAlerts = False

            wkbook.Close

            Application.DisplayAlerts = True

        End If

End If

'clear and update the related items in listbox
lstOWB.Clear
lstBook.Clear
'cmbForecast.Clear

For intlist = 1 To Workbooks.count
    
    If Workbooks(intlist).Name <> ThisWorkbook.Name Then
        lstOWB.AddItem Workbooks(intlist).Name
        lstBook.AddItem Workbooks(intlist).Name

        'cmbForecast.AddItem Workbooks(intlist).Name
    End If
    
Next intlist

End Sub

Private Sub btnCreate_Click()

Dim intlist As Integer, book_name As String, sheet As Worksheet, available As Integer

'check if the uesr selected a workbook
If wbSum.ListCount < 1 Then

    MsgBox "Please select workbooks for creating summary."

    Exit Sub

End If

'check if the reuquired worksheet and records are exist in the workbook
For intlist = 0 To wbSum.ListCount - 1

    For Each sheet In Workbooks(wbSum.List(intlist)).Worksheets

        If sheet.Name = "Monthly Summary" Then

            available = 1

            Exit For

        Else

            available = 0

        End If

    Next sheet

'if the required worksheet is not exist, the message will be shown
    If available = 0 Then

        MsgBox "The Monthly Summary is not available in " & wbSum.List(intlist) & "." & vbCrLf & "Please generate a Monthly Summary in previous page."

        Exit Sub

    End If

Next intlist

'create a new workbook with name according to option
Workbooks.Add

If btnYear = True Then

    book_name = "Year_Report"

Else

    book_name = "Month_Report"

End If


Set Summary_Book = ActiveWorkbook

Summary_Book.Windows(1).Caption = book_name
Summary_Book.Windows(1).Visible = False

Dim ref_wkbook As Workbook, ref_wksheet As Worksheet
Dim ref_row As Long, no_month As Integer, total_price As Double, ref_col As Long, summary_row As Long, total_sale As Double, total_profit As Double


'create a worksheet in the new workbook with the product name
For intlist = 0 To wbSum.ListCount - 1

    Set ref_wkbook = Workbooks(wbSum.List(intlist))
    Set Summary_Sheet = Summary_Book.Sheets.Add
    
'remove the characters ".xlsx"
    Summary_Sheet.Name = VBA.Left(wbSum.List(intlist), VBA.Len(wbSum.List(intlist)) - 5)

    Set ref_wksheet = ref_wkbook.Sheets("Monthly Summary")

'the follwing code is used to generate the yearly summary
    If btnYear = True Then

'insert the header to the new worksheet
        With Summary_Sheet
        .Range("a1").Value = "Year"
        .Range("b1").Value = "Average Profit Among Item Types"
        .Range("c1").Value = "Total Profit"
        '.Range("d1").Value = "Total Sale Among Regions (10 Million)"
        End With

'extract the distinct year value from the dataset
        For ref_row = 2 To ref_wksheet.Cells(Rows.count, 1).End(xlUp).row

            If ref_wksheet.Cells(ref_row, 1) <> ref_wksheet.Cells(ref_row + 1, 1) Then

            Summary_Sheet.Range("a1048576").End(xlUp).Offset(1, 0).Value = ref_wksheet.Cells(ref_row, 1).Value

            End If

        Next ref_row

'accumulate the the variables
        For summary_row = 2 To Summary_Sheet.Cells(Rows.count, 1).End(xlUp).row

            no_month = 0
            total_price = 0
            total_sale = 0
            total_profit = 0



'match the year by row
                For ref_row = 2 To ref_wksheet.Cells(Rows.count, 1).End(xlUp).row

                    If VBA.Val(ref_wksheet.Cells(ref_row, 1).Value) = VBA.Val(Summary_Sheet.Cells(summary_row, 1).Value) Then

                        no_month = no_month + 1

'accumulate the the value for specific columns
                    For ref_col = 1 To ref_wksheet.Cells(1, Columns.count).End(xlToLeft).Column


                    If ref_wksheet.Cells(1, ref_col).Value Like "Total Profit_[1-12]" Then

                        total_profit = total_profit + VBA.Val(ref_wksheet.Cells(ref_row, ref_col).Value)

                    End If

                        Next ref_col

                    End If



            Next ref_row
            
'output final result when it reach the last specific year
Summary_Sheet.Range("b1048576").End(xlUp).Offset(1, 0).Value = total_profit / (no_month * 12)
Summary_Sheet.Range("c1048576").End(xlUp).Offset(1, 0).Value = total_profit
'Summary_Sheet.Range("d1048576").End(xlUp).Offset(1, 0).Value = total_sale / 10000000

Next summary_row

End If


'the follwing code is used to generate the monthly summary
If btnMonth = True Then

'insert the header in the new worksheet
    With Summary_Sheet
    .Range("a1").Value = "Year"
    .Range("b1").Value = "Month"
    .Range("c1").Value = "Average Profit Per Months"
    .Range("d1").Value = "Average Profit Among Item Types"
    .Range("e1").Value = "Total Profit"
    End With

    For ref_row = 2 To ref_wksheet.Cells(Rows.count, 1).End(xlUp).row

        'total_price = 0
        'total_sale = 0
        total_profit = 0


        Summary_Sheet.Range("a1048576").End(xlUp).Offset(1, 0).Value = ref_wksheet.Cells(ref_row, 1).Value
        Summary_Sheet.Range("b1048576").End(xlUp).Offset(1, 0).Value = ref_wksheet.Cells(ref_row, 2).Value

'since the monthly summary contains the price and sale data for each month on a year
'sum up the price and sale data for each row
        For ref_col = 1 To ref_wksheet.Cells(1, Columns.count).End(xlToLeft).Column

            'If ref_wksheet.Cells(1, ref_col).Value Like "Avg_Price_[1-5]" Then

                 'total_price = total_price + VBA.Val(ref_wksheet.Cells(ref_row, ref_col).Value)

            'End If

            If ref_wksheet.Cells(1, ref_col).Value Like "Total Profit_[1-12]" Then

            total_profit = total_profit + VBA.Val(ref_wksheet.Cells(ref_row, ref_col).Value)

            End If

            Next ref_col

'output the final result
    Summary_Sheet.Range("c1048576").End(xlUp).Offset(1, 0).Value = total_profit / 12
    Summary_Sheet.Range("d1048576").End(xlUp).Offset(1, 0).Value = total_profit / 12
    Summary_Sheet.Range("e1048576").End(xlUp).Offset(1, 0).Value = total_profit

    Next ref_row


    End If

Next intlist

'delete the sheet1 which is useless
Application.DisplayAlerts = False
Summary_Book.Worksheets("sheet1").Delete
Application.DisplayAlerts = True

frmReport.Visible = True

End Sub

Private Sub btnDelSheet_Click()

If lstWS.ListIndex = -1 Then

    MsgBox "Select a worksheet to proceed."
    
Else

    If wkbook.Worksheets.count = 1 Then
    
        MsgBox "Cannot delete the last worksheet."
        
        Exit Sub
        
    Else
    
        Application.DisplayAlerts = False
        
        If MsgBox("Confirm your delete action!", vbYesNo, "Confirm") = vbYes Then
        
            wksheet.Delete
            
        End If
        
        Application.DisplayAlerts = True
        
        lstWS.Clear
        
        For Each wksheet In wkbook.Worksheets
        
            lstWS.AddItem wksheet.Name
            
        Next wksheet
    
    End If
    
End If

End Sub

Private Sub btnDelRec_Click()

Dim row As Long, counter As Integer, count As Integer

For count = 0 To lstWS.ListCount - 1

    If lstWS.Selected(count) = True Then
    
        If lstWS.List(count) = "Monthly Summary" Then
        
            MsgBox ("Month Summary are not allowed to delete" & VBA.vbCrLf & "Please edit the weekly data in order to update the monthly summary report.")
            
            Exit Sub
            
        End If
        
    End If
    
Next count

If lstRecord.ListIndex = -1 Then

    MsgBox "Please select a record!"
    
Else

    If MsgBox("Confirm your delete action!", vbYesNo, "Confirm") = vbYes Then
    
        Rows(lstRecord.ListIndex + 1).Delete
        
    End If
    
End If


End Sub

Private Sub btnInsert_Click()

Data_Change.Insert_Update.Caption = "Insert"

Dim counter As Integer

If lstWS.ListIndex = -1 Then

    MsgBox "Please select a worksheet!"
    
    Exit Sub
    
End If

If wksheet.Name = "Monthly Summary" Then

    MsgBox ("Month Summary are not allowed to insert record!" & VBA.vbCrLf & "Please edit the weekly data in order to update the Monthly Summary.")
    
Else

    For counter = 1 To 8
    
        Data_Change.Controls("Label" & counter).Caption = wksheet.Range("a1").Offset(0, counter - 1)
    
    Next counter
    
    For counter = 1 To 8
    
        If Data_Change.Controls("Label" & counter).Caption = "" Then
        
            Data_Change.Controls("TextBox" & counter).Visible = False
            
        End If
        
    Next counter
    
    Data_Change.TextBox2.Value = Month(Now)
    
    Data_Change.TextBox1.Value = Year(Now)
    
    Data_Change.month_change.Value = Month(Now)
    
    Data_Change.Show
    
End If

End Sub

Private Sub btnInsertWB_Click()

Dim intlist As Integer

'combine all selected workbooks
For intlist = lstBook.ListCount - 1 To 0 Step -1

    If lstBook.Selected(intlist) = True Then
    
    wbSum.AddItem lstBook.List(intlist)
       
    lstBook.RemoveItem intlist
    
    End If
    
Next intlist

End Sub

Private Sub btnList_Click()

lstDirWB.Clear

Dim strFilename As String

strDir = textDir.Value
'strDir = "C:\Users\Asus\Desktop\Calvin\VBA Project\"

strFilename = VBA.Dir(strDir & "*.xlsx")

If strFilename = "" Then
    
    VBA.MsgBox "Either invalid directory or the directory does not contain" & "any Excel file."
    
    Exit Sub
    
End If

Do Until strFilename = ""
    
    lstDirWB.AddItem strFilename
    
    strFilename = VBA.Dir
    
Loop


End Sub

Private Sub btnModify_Click()

Data_Change.Insert_Update.Caption = "Update"

Dim counter As Integer, listitem As Long

If lstRecord.ListIndex = -1 Then

    MsgBox "Please select a record!"
    
    Exit Sub
    
End If

If wksheet.Name = "Monthly Summary" Then

    MsgBox ("Month Summary are not allowed to modify record!" & VBA.vbCrLf & "Please edit the weekly data in order to update the Monthly Summary.")
    
Else

    For listitem = 0 To lstRecord.ListCount - 1
    
        If lstRecord.Selected(listitem) = True Then
        
            For counter = 1 To wksheet.Cells(1, wksheet.Columns.count).End(xlToLeft).Column
            
                Data_Change.Controls("Label" & counter).Caption = wksheet.Range("a1").Offset(0, counter - 1)
            
                Data_Change.Controls("TextBox" & counter).Text = wksheet.Cells(listitem + 1, 1).Offset(0, counter - 1)
            
            Next counter
            
            Exit For
            
        End If
        
    Next listitem
    
    For counter = 1 To 8
    
        If Data_Change.Controls("Label" & counter).Caption = "" Then
        
            Data_Change.Controls("TextBox" & counter).Visible = False
            
        End If
        
    Next counter
    
    Data_Change.month_change.Value = Data_Change.TextBox1.Value
    
    Data_Change.Show
    
End If
    

End Sub

Private Sub btnMonth_Click()

End Sub

Private Sub btnMonthReport_Change()

viewSummary.Clear

Dim book As Workbook, sheet As Worksheet

If btnMonthReport = True Then

    For Each book In Workbooks

        If book.Windows(1).Caption = "Month_Report" Then
        book.Activate

            For Each sheet In book.Worksheets

                viewSummary.AddItem sheet.Name

            Next sheet

        End If

    Next book

End If

End Sub

Private Sub btnPivot_Click()
    
 ' Creates a chart based on a PivotTable report.
    Dim objPivot As PivotTable, objPivotRange As Range, objChart As Chart
    
    ' Call the CreatePivot macro to create a new PivotTable report.
    CreatePivot
    
    ' Access the new PivotTable from the sheet's PivotTables collection.
    Set objPivot = ActiveSheet.PivotTables(1)
    
    ' Add a new chart sheet.
    Set objChart = Charts.Add
    
    ' Create a Range object that contains
    ' all of the PivotTable data, except the page fields.
    Set objPivotRange = objPivot.TableRange1
    
' Specify the PivotTable data as the chart's source data.
    With objChart
        .SetSourceData objPivotRange
        .ChartType = xl3DColumn
        .Legend.Delete
    End With
    
    ActiveWorkbook.Windows(1).Visible = True
    

End Sub

Sub CreatePivot()

Dim source As Range, pivot_sheet As Worksheet

Set source = wksheet.Range("a1", wksheet.Range("a1").End(xlDown).End(xlToRight))
Set pivot_sheet = Sheets.Add
pivot_sheet.Activate
pivot_sheet.Name = "PivotTable"

ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        source, Version:=6).CreatePivotTable TableDestination _
        :="PivotTable!R1C1", TableName:="PivotTable2", DefaultVersion:=6
    Sheets("PivotTable").Select
    Cells(1, 1).Select
    With ActiveSheet.PivotTables("PivotTable2")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    
ActiveSheet.PivotTables("PivotTable2").RepeatAllLabels xlRepeatLabels
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData source:=Range("PivotTable!$A$1:$C$36")
    ActiveSheet.Shapes("Chart 1").IncrementLeft 192
    ActiveSheet.Shapes("Chart 1").IncrementTop 15
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Year")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveChart.PivotLayout.PivotTable.PivotFields("Month")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Total Profit_1"), "Baby Food", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Total Profit_2"), "Beverages", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Total Profit_3"), "Cereal", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Total Profit_4"), "Clothes", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Total Profit_5"), "Cosmetics", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Total Profit_6"), "Fruits", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Total Profit_7"), "Household", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Total Profit_8"), "Meat", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Total Profit_9"), "Office Supplies", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Total Profit_10"), "Personal Care", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Total Profit_11"), "Snacks", xlSum
    ActiveChart.PivotLayout.PivotTable.AddDataField ActiveChart.PivotLayout. _
        PivotTable.PivotFields("Total Profit_12"), "Vegetables", xlSum
 
    ActiveSheet.Shapes("Chart 1").IncrementLeft -294
    ActiveSheet.Shapes("Chart 1").IncrementTop 148.5
    ActiveChart.ShowValueFieldButtons = False

End Sub


Private Sub btnSave_Click()

Dim sheet As Integer

'check if the user selected a workbook
If wkbook Is Nothing Then

    MsgBox "Please select a workbook!"

    Else

    wkbook.Save

    MsgBox "The selected workbook is saved successfully!"

    End If

'consert and save the dynamic pivot table to static table which the user last seen
'For sheet = 1 To wkbook.Sheets.count

    'If wkbook.Worksheets(sheet).Name = "PivotTable" Then
    'If wkbook.Sheets(sheet).Name = "PivotTable" Then

    'wkbook.Sheets(sheet).Range("a2").CurrentRegion.Select

    'Selection.Copy
    'wkbook.Sheets(sheet).Range("A1048576").End(xlUp).Offset(1, 0).Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    'End If
'Next sheet

End Sub

Private Sub btnSaveSummary_Click()

Summary_Book.SaveAs Filename:=Application.GetSaveAsFilename(filefilter:="excel file,*.xlsx")

End Sub

Private Sub btnSummary_Click()

Dim count_sheet As Integer, region As Integer, row As Long, ref_sheet As Worksheet

'message displayed if user do not select a workbook
If lstOWB.ListIndex = -1 Then

    MsgBox "Please select a Workbook!"

    Exit Sub

    Else

'check if the monthly summary is aready exist
    For count_sheet = 0 To lstWS.ListCount - 1

        If VBA.LCase(lstWS.List(count_sheet)) = "Monthly Summary" Then

        MsgBox "The Monthly Summary already exists."
        
'ask for updating if the monthly summary is existed in the workbook
            If MsgBox("Do you want to update the Monthly Summary?", vbYesNo, "Confirm") = vbYes Then
 
            Application.DisplayAlerts = False
            
'delete the original one
            wkbook.Worksheets("Monthly Summary").Delete

            lstWS.RemoveItem count_sheet

            Application.DisplayAlerts = True
            
            Else

            Exit Sub
 
            End If

        Exit For

        End If

    Next count_sheet

End If

'check if the required data is contained in the workbook
For count_sheet = 0 To lstWS.ListCount - 1
 
    If lstWS.List(count_sheet) = "Sheet1" Then
        
    Set ref_sheet = wkbook.Worksheets("Sheet1")
    
    Exit For
        
    End If

Next count_sheet

'message for the case if the required data ommited
If ref_sheet Is Nothing Then

    MsgBox "The Monthly Summary can not be created because of the shortage of data!"

    Exit Sub

End If

'create new worksheet and assign a name
Worksheets.Add(After:=Sheets(Sheets.count)).Name = "Monthly Summary"

lstWS.AddItem "Monthly Summary"

wkbook.Worksheets("Monthly Summary").Activate

'insert the header for the new worksheet
For region = 1 To 12

    Cells(1, 2 + region).Value = "Total Profit_" & region
    'Cells(1, 2 + 2 * region).Value = "Total_Sale_" & region

Next region

'copy the first two columns data in the weekly workbook
ref_sheet.Range("a:a", "b:b").Copy Destination:=Range("a1")

'delete the duplicate
'For row = Cells(Rows.count, 1).End(xlUp).row To 2 Step -1

    'If Cells(row, 2).Value = Cells(row - 1, 2).Value Then

    'Rows(row - 1).Delete

    'End If

'Next row

For row_1 = Cells(Rows.count, 1).End(xlUp).row To 3 Step -1

    For row_2 = row_1 - 1 To 2 Step -1
    
    If Cells(row_1, 1).Value = Cells(row_2, 1).Value And Cells(row_1, 2).Value = Cells(row_2, 2).Value Then
    
    Rows(row_2).Delete
    
    End If
    
    Next row_2
Next row_1


Dim col_profit As Long, col_item As Long

'find the price, sale and region columns in the weekly data
'col_price = Application.WorksheetFunction.Match("Weighted Price", ref_sheet.Range("a1", "XFD1"), 0)

col_profit = Application.WorksheetFunction.Match("Total Profit", ref_sheet.Range("a1", "XFD1"), 0)

col_item = Application.WorksheetFunction.Match("Item Type", ref_sheet.Range("a1", "XFD1"), 0)

'accumulate and calculate the sale and price according to the region, yesr and month
Dim total_profit As Double, ref_row As Long

For region = 1 To 12

    For row = 2 To Cells(Rows.count, 1).End(xlUp).row

    'total_sale = 0
    total_profit = 0

        For ref_row = 2 To ref_sheet.Cells(Rows.count, 1).End(xlUp).row

'match the year, month and data
        If ref_sheet.Cells(ref_row, 1).Value = Cells(row, 1).Value And _
        ref_sheet.Cells(ref_row, 2).Value = Cells(row, 2).Value And _
        ref_sheet.Cells(ref_row, col_item).Value = region Then
   
'accumulate the total
        'total_sale = total_sale + VBA.Val(ref_sheet.Cells(ref_row, col_sale).Value)
        total_profit = total_profit + VBA.Val(ref_sheet.Cells(ref_row, col_profit).Value)

        End If

        Next ref_row
        
'output the result
    Cells(row, 2 + region).Value = total_profit

    Next row

Next region



End Sub


Private Sub btnOpen_Click()

Dim intlist As Integer

'strDir = "C:\Users\Asus\Desktop\Calvin\VBA Project\"

For intlist = 0 To lstDirWB.ListCount - 1

    If lstDirWB.Selected(intlist) = True Then
    
        Workbooks.Open Filename:=strDir & lstDirWB.List(intlist)
    
    Workbooks(lstDirWB.List(intlist)).Windows(1).Visible = False
    
    End If
    
Next intlist

lstDirWB.Clear

For intlist = 1 To Workbooks.count

    If Workbooks(intlist).Name <> ThisWorkbook.Name Then
    
        lstOWB.AddItem Workbooks(intlist).Name
    
        lstBook.AddItem Workbooks(intlist).Name
    
        'cmbForecast.AddItem Workbooks(intlist).Name
    End If
    
Next intlist

For intlist = lstDirWB.ListCount - 1 To 0 Step -1

    If lstDirWB.Selected(intlist) = True Then
    
        lstDirWB.RemoveItem intlist
    
    End If
    
Next intlist

End Sub


Private Sub btnDrop_Click()

Dim intlist As Integer

'transfer the items between listboxes from right to left
For intlist = wbSum.ListCount - 1 To 0 Step -1

    If wbSum.Selected(intlist) = True Then
        lstBook.AddItem wbSum.List(intlist)
        wbSum.RemoveItem intlist
    End If
    
Next intlist

End Sub

Private Sub btnView_Click()

Dim sheet As Worksheet

Set sheet = ActiveWorkbook.Worksheets(viewSummary.Value)

sheet.Activate

lstReport.RowSource = ""

With lstReport
.ColumnCount = sheet.Cells(1, sheet.Columns.count).End(xlToLeft).Column
.RowSource = sheet.Range("a1", sheet.Range("A1").End(xlToRight).End(xlDown)).Address
End With


Dim chrtEmbeddObj As Shape, chrtEmbedd As Chart

If btnYearReport = True Then

'Add an embedded Shape object on wksheet and assign it to chrtEmbeddObj
Set chrtEmbeddObj = sheet.Shapes.AddChart2(XlChartType:=xlPie, Width:=402, Height:=192)

'Assign the embedded chart object to chrtEmbedd
Set chrtEmbedd = chrtEmbeddObj.Chart

'Assign a name to the shape
chrtEmbeddObj.Name = "Chart "
    
'Specify the data source
chrtEmbedd.SeriesCollection.NewSeries

With sheet
    chrtEmbedd.SeriesCollection(1).XValues = .Range(.Range("a2"), .Range("a2").End(xlDown).Offset(1, 0))
    chrtEmbedd.SeriesCollection(1).Values = .Range(.Range("c2"), .Range("c2").End(xlDown).Offset(1, 0))

End With

'Specify the layout of the chart
chrtEmbedd.ApplyLayout (1)

'Specify the title of the chart
chrtEmbedd.SetElement element:=msoElementChartTitleAboveChart
chrtEmbedd.ChartTitle.Caption = "Yearly Sale"

End If

If btnMonthReport = True Then

Dim rngSource As Range, rngBin As Range, count As Integer, bin As Long, location As String

'create the bin value for the data (5 classes)
bin = 0

sheet.Range("B1").Value = "Profit"

For count = 1 To 6
    sheet.Range("L1").Offset(count, 0).Value = bin
    bin = bin + (Application.WorksheetFunction.Max(sheet.Range("D1:D1048576")) + Application.WorksheetFunction.Min(sheet.Range("D1:d1048576"))) / 5
Next count

'create the histogram for the profit
With sheet


    'Define the data series including the column header
    Set rngSource = .Range("d1", .Range("d1048576").End(xlUp))
    
    'Define the bin series including the column header
    Set rngBin = .Range("L1", .Range("L1048576").End(xlUp))

    Application.Run "ATPVBAEN.XLAM!Histogram", rngSource, _
        .Range("$n$1"), rngBin, False, False, True, True

    'Cannot use ActiveChart object to refer to the histogram as it is not activated
    Set chrtEmbedd = sheet.Shapes(sheet.Shapes.count).Chart

    .Range("L1", .Range("L1").End(xlDown).Address).Clear
End With

With chrtEmbedd

    'Set the gapwidth of the bars to 0% of the bar width
    .ChartGroups(1).GapWidth = 0

    'Turn off the legend
    .SetElement element:=msoElementLegendNone
    
    'Define the chart title text
    .ChartTitle.Caption = "Profit"
  

End With


End If

'expoert the created chart and load into the userform for precview
Dim image_name As String

image_name = Application.DefaultFilePath & Application.PathSeparator & "temp.gif"
chrtEmbedd.Export Filename:=image_name

Application.ScreenUpdating = True

displayChart.Picture = LoadPicture(image_name)

End Sub

Private Sub btnYear_Click()

End Sub

Private Sub btnYearReport_Change()

viewSummary.Clear

Dim book As Workbook, sheet As Worksheet

If btnYearReport = True Then

    For Each book In Workbooks

        If book.Windows(1).Caption = "Year_Report" Then

            book.Activate

            For Each sheet In book.Worksheets

                viewSummary.AddItem sheet.Name

            Next sheet

        End If

    Next book

End If

End Sub



Private Sub lstOWB_Change()

Dim counter As Integer, intlist As Integer

lstWS.Clear

For counter = 0 To lstOWB.ListCount - 1

    If lstOWB.Selected(counter) = True Then
    
    btnClose.Enabled = True
    
    btnSave.Enabled = True
    
    btnDelSheet.Enabled = True
    
    btnSummary.Enabled = True
    
    MultiPage1.Pages(1).Enabled = True
    
    'MultiPage1.Pages(2).Enabled = True
    
    Set wkbook = Workbooks(lstOWB.List(counter))
    
    wkbook.Activate
    
    For Each wksheet In wkbook.Worksheets
    
    lstWS.AddItem wksheet.Name
    
    Next wksheet
    
    Exit For
    
    Else
    
    btnClose.Enabled = False
    
    btnSave.Enabled = False
    
    btnDelSheet.Enabled = False
    
    btnSummary.Enabled = False
    
    End If
    
Next counter

For intlist = 0 To lstWS.ListCount - 1

    lstWS.Selected(intlist) = False
    
Next intlist

End Sub

Private Sub lstRecord_Click()

End Sub

Private Sub lstDirWB_Change()

Dim count As Integer

For count = 0 To lstDirWB.Selected(count) - 1

    If lstDirWB.Selected(count) = True Then
    
    btnOpen.Enabled = True
    
    Exit For
    
    Else
    
    btnOpen.Enabled = False
    
    End If
    
Next count

End Sub

Private Sub lstWS_Change()

lstRecord.RowSource = ""

Dim counter_sheet As Integer, counter_book As Integer

For counter_sheet = 0 To lstWS.ListCount - 1

    If lstWS.Selected(counter_sheet) = True Then
        
        Set wksheet = wkbook.Worksheets(lstWS.List(counter_sheet))
        
        wksheet.Activate
        
        With lstRecord
        
        .ColumnCount = wksheet.Cells(1, wksheet.Columns.count).End(xlToLeft).Column
        
        .RowSource = wksheet.Range("a1", wksheet.Range("H1").End(xlDown)).Address
        
        End With
        
        'If Worksheets(lstWorksheet.List(counter_sheet)).Name = "Yearly Summary" Then
        If Worksheets(lstWS.List(counter_sheet)).Name = "Monthly Summary" Then
        
        btnPivot.Enabled = True
        
        Else
        
        btnPivot.Enabled = False
        
        End If
        
    Exit For
    
    End If
    
Next counter_sheet

End Sub
