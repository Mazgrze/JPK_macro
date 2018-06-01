Sub JPK_Creator()

' Remember start epoch for benchmarking
Dim start_time As Long
start_time = DateDiff("s", #1/1/1970#, Now())

' Disable auto updating for performance
Application.ScreenUpdating = False
Application.Calculation = xlManual

' Declaring names of sheets and pivot table
Const mapping_sheet = "Mapping" '
Const jpk_sheet = "JPK"
Const pvt_sheet = "pivot_summary"
Const piv_name = "Summary pivot"

Const taxCode = 1
Const invoice_no = 2
Const contractor_address = 3
Const contractor_no = 4
Const contractor_name = 5
Const document_date = 6
Const transaction_date = 7
Const origin_amount = 8
Const tax_amount = 9



' ////////////////////////////////////////////Checking sheets before start
If Not Check_worksheet(mapping_sheet) Then
    MsgBox ("There is no " & mapping_sheet & " sheet. Create or rename sheet with tax code mappings.")
    Exit Sub
End If

If LCase(Worksheets(mapping_sheet).Cells(1, 1).Text) <> "tax code" Or _
    LCase(Worksheets(mapping_sheet).Cells(1, 2).Text) <> "amount" Then
        MsgBox ("Incorrect mapping table or " & mapping_sheet & " sheet not found.")
        Exit Sub
End If

If Worksheets(1).ListObjects.Count = 0 Then
    'When there is no table in first sheet
    MsgBox ("Data sheet must be formatted as a Table (Ctrl + L).")
    Exit Sub
End If

' Check headers in found table
Dim headers
headers = Worksheets(1).ListObjects(1).HeaderRowRange
If UBound(headers, 2) < 10 Then
    ' Found wrong table
    MsgBox ("Found incorrect table. Make sure sheet with data is first in cards (leftmost).")
    Exit Sub
End If

' Check if headers contain necessary columns with proper names
Dim required, col_err As Boolean
col_err = False

' Create array with needed column names
required = Array("Sales tax code", "Invoice", "Address", "Tax exempt number", _
            "Customer/Vendor", "Document date", "Date", "Amount origin", "Sales tax amount")

For i = LBound(required, 1) To UBound(required, 1)
    For j = 1 To UBound(headers, 2)
    If required(i) = headers(1, j) Then
        Exit For
    End If
    If j = UBound(headers, 2) Then
        col_err = True
        ' Alert about lack of column
        MsgBox ("`" & required(i) & "` column not found. Please rename the column.")
    End If
    Next j
Next i
If col_err Then Exit Sub ' After all alerts exit sub


Dim source_table_name As String
source_table_name = Worksheets(1).ListObjects(1).name

' Create new sheet for pivot table and create pivot table
Make_sheet (pvt_sheet)
ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        source_table_name, Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:=Worksheets(pvt_sheet).Range("A1"), _
        TableName:=piv_name, DefaultVersion:=xlPivotTableVersion14



' Assign pivot to object
Dim pivot As PivotTable
Set pivot = Sheets(pvt_sheet).PivotTables(piv_name)

' Format pivot as table
With pivot
    .RowAxisLayout xlTabularRow
    .ShowTableStyleRowHeaders = False
    .ShowTableStyleColumnHeaders = True
    .RepeatAllLabels xlRepeatLabels
    .ColumnGrand = False
    .RowGrand = False
End With

' Remove subtotals from pivot
For i = 1 To ActiveSheet.PivotTables(1).PivotFields.Count - 1
     ActiveSheet.PivotTables(1).PivotFields(i).Subtotals(1) = False
Next i

' //////////////////////////////////////  Adding pivot fields
For i = LBound(required, 1) To (UBound(required, 1) - 2)
    pivot.PivotFields(required(i)).Orientation = xlRowField
Next i

pivot.AddDataField pivot.PivotFields(required(7)), "Sum of " & required(7), xlSum
pivot.AddDataField pivot.PivotFields(required(8)), "Sum of " & required(8), xlSum


Dim data_range, map_range, entries_number As Integer

' Read entries from pivot into array for performance. First row in range is the headings row.
data_range = ActiveSheet.PivotTables(1).TableRange1

' Count rows in pivot. Value is decremented by 1 because of headings row in first array element
entries_number = UBound(data_range, 1) - 1

' Check if there are any entries in pivot
If entries_number < 2 And (data_range(2, 1) Like "([a-z]*[a-z])" Or data_range(2, 1) = "") Then
    MsgBox ("Table " & Sheets(1).ListObjects(1).name & " in sheet " & Sheets(1).name & " is empty")
    Exit Sub
End If

' Declare arrays containers for sales and purcharse entries 
Dim purcharses()
Dim sales()

' Some debug feedback for control
Debug.Print "Found: " & entries_number & " Entries"

' Reserve memory for arrays
ReDim sales(1 To entries_number, 1 To 37)
ReDim purcharses(1 To entries_number, 1 To 19)

' Declarations of variables needed for main loop
Dim current_code As String ' variable for current tax code
Dim column_numer As Integer, jpk_row As Long, input_row As Long
Dim sales_counter As Long, purcharses_counter As Long

Dim sales_tax_sum As Double, purch_tax_sum As Double ' containers for summing taxes
Dim map_rows As Integer ' number of tax codes
Dim mapped_col(), mapped_amount() ' mapped column and mapped amount arrays
Dim map_nr As Integer, sales_mappings As Integer, purch_mappings As Integer ' counters for mapping
Dim company_info, company_info2, columns ' Arrays with company info

sales_tax_sum = 0
purch_tax_sum = 0
input_row = 1
jpk_row = 1
purcharses_counter = 1
sales_counter = 1
column_numer = 1

'/// COLUMNS NAMES /////////////////////////////////////////////
columns = Array("KodFormularza", "kodSystemowy", "wersjaSchemy", "WariantFormularza", _
                "CelZlozenia", "DataWytworzeniaJPK", "DataOd", "DataDo", "NazwaSystemu", _
                "NIP", "PelnaNazwa", "Email", "LpSprzedazy", "NrKontrahenta", "NazwaKontrahenta", _
                "AdresKontrahenta", "DowodSprzedazy", "DataWystawienia", "DataSprzedazy", _
                "K_10", "K_11", "K_12", "K_13", "K_14", "K_15", "K_16", "K_17", "K_18", "K_19", _
                "K_20", "K_21", "K_22", "K_23", "K_24", "K_25", "K_26", "K_27", "K_28", "K_29", _
                "K_30", "K_31", "K_32", "K_33", "K_34", "K_35", "K_36", "K_37", "K_38", "K_39", _
                "LiczbaWierszySprzedazy", "PodatekNalezny", "LpZakupu", "NrDostawcy", "NazwaDostawcy", _
                "AdresDostawcy", "DowodZakupu", "DataZakupu", "DataWplywu", "K_43", "K_44", "K_45", _
                "K_46", "K_47", "K_48", "K_49", "K_50", "LiczbaWierszyZakupow", "PodatekNaliczony")
                
company_info = Array("JPK_VAT", "JPK_VAT (3)", "1-1", "3", "0", "", "", "", "SOFTWARE NAME")
company_info2 = Array("COMPANY VAT NUMBER", "COMPANY NAME", "")

' Write current Date/time into company info array
company_info(5) = Format(Date, "yyyy") & "-" & Format(Date, "mm") & "-" & Format(Date, "dd") _
                & "T" & Format(Time, "hh") & ":" & Format(Time, "nn") & ":" & Format(Time, "ss")
				
' Write accounting period start date (First day of previous month)
company_info(6) = "01." & Format(Date - Format(Date, "dd") - 10, "mm") & "." & Format(Date, "yyyy")

' Write accounting period end date (last day of previous month)
company_info(7) = MonthDays(Int(Format(Date - -Format(Date, "dd") - 10, "m"))) & "." & _
                Format(Date - Format(Date, "dd") - 10, "mm") & "." & Format(Date, "yyyy")
		
' Create Final Sheet
Make_sheet (jpk_sheet)

' //////////////////////////////// Write column names and company info into sheet
For Each col In columns
    Worksheets(jpk_sheet).Cells(jpk_row, column_numer).NumberFormat = "@"
    Worksheets(jpk_sheet).Cells(jpk_row, column_numer).Value = col
    column_numer = column_numer + 1
Next col

column_numer = 1
jpk_row = jpk_row + 1
For Each col In company_info
    Worksheets(jpk_sheet).Cells(jpk_row, column_numer).NumberFormat = "@"
    Worksheets(jpk_sheet).Cells(jpk_row, column_numer).Value = col
    column_numer = column_numer + 1
Next col

jpk_row = jpk_row + 1
For Each col In company_info2
    Worksheets(jpk_sheet).Cells(jpk_row, column_numer).NumberFormat = "@"
    Worksheets(jpk_sheet).Cells(jpk_row, column_numer).Value = col
    column_numer = column_numer + 1
Next col
jpk_row = jpk_row + 1


' Read mapping table into array. First recognise size of table and then read into array
map_rows = Worksheets(mapping_sheet).Range("a900000").End(xlUp).row
map_range = Worksheets(mapping_sheet).Range(Worksheets(mapping_sheet).Cells(2, 1), _
                Worksheets(mapping_sheet).Cells(map_rows, 3))


' MAIN LOOP START //////////////////////////////////////////////////////////////////
For i = 1 To entries_number

    purch_mappings = 0
    sales_mappings = 0
	
'\\\\\\\\\\\\\\\ EVERYWERE IN CODE DATA_RANGE HAS "i+1" ITERATION VALUE BECAUSE OF HEADINGS IN FIRST ELEMENT \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    current_code = Trim(UCase(data_range(i + 1, taxCode)))
    map_nr = 1
    
	
    For j = 1 To UBound(map_range, 1)
        If Trim(UCase(map_range(j, 1))) = current_code Then
            ReDim Preserve mapped_col(map_nr)
            ReDim Preserve mapped_amount(map_nr)
            mapped_col(map_nr) = map_range(j, 3)
            If map_range(j, 3) < 40 Then
				sales_mappings = sales_mappings + 1
            Else
				purch_mappings = purch_mappings + 1
            End If
            
			If Trim(LCase(map_range(j, 2))) = "origin" Then
                mapped_amount(map_nr) = data_range(i + 1, origin_amount)
            Else
                mapped_amount(map_nr) = data_range(i + 1, tax_amount)
            End If
            map_nr = map_nr + 1
        End If
    Next j
    
    
    If purch_mappings Then
         purcharses(purcharses_counter, 1) = purcharses_counter
         purcharses(purcharses_counter, 2) = data_range(input_row + 1, contractor_no)
         purcharses(purcharses_counter, 3) = data_range(input_row + 1, contractor_name)
         purcharses(purcharses_counter, 4) = data_range(input_row + 1, contractor_address)
         purcharses(purcharses_counter, 5) = data_range(input_row + 1, invoice_no)
         purcharses(purcharses_counter, 6) = data_range(input_row + 1, transaction_date)
         purcharses(purcharses_counter, 7) = data_range(input_row + 1, document_date)
         
         For k = 8 To 15
            purcharses(purcharses_counter, k) = 0
         Next k
         
        For n = 1 To UBound(mapped_col, 1)
            If mapped_col(n) > 39 Then
                purcharses(purcharses_counter, mapped_col(n) - 35) = mapped_amount(n)
            End If
        Next n
        
        purcharses_counter = purcharses_counter + 1
    End If
    
    
    If sales_mappings Then
            sales(sales_counter, 1) = sales_counter
            sales(sales_counter, 2) = data_range(input_row + 1, contractor_no)
            sales(sales_counter, 3) = data_range(input_row + 1, contractor_name)
            sales(sales_counter, 4) = data_range(input_row + 1, contractor_address)
            sales(sales_counter, 5) = data_range(input_row + 1, invoice_no)
            sales(sales_counter, 6) = data_range(input_row + 1, transaction_date)
            sales(sales_counter, 7) = data_range(input_row + 1, document_date)
            For k = 8 To 37
               sales(sales_counter, k) = 0
            Next k
            
            Dim tekst
            tekst = ""
            For n = 1 To UBound(mapped_col, 1)
                If mapped_col(n) < 40 Then
                    sales(sales_counter, mapped_col(n) - 2) = mapped_amount(n)
                    tekst = tekst & " " & mapped_col(n)
                End If
            Next n
            
            sales_counter = sales_counter + 1
            data_range(input_row + 1, invoice_no) = "sale " & tekst
    End If
      
    input_row = input_row + 1

Next i

Dim suma

' Print size of created arrays for control
Debug.Print "final Sales: " & sales_counter & _
            "  final Purch.: " & purcharses_counter


' Summing loops
sales_tax_sum = 0
For i = 16 To 37
    suma = 0
    If i = 16 Or i = 18 Or i = 20 Or i = 24 Or i = 26 Or i = 28 _
        Or i = 30 Or i = 33 Or i = 35 Or i = 36 Or i = 37 Then
        For j = LBound(sales, 1) To UBound(sales, 1)
            suma = suma + sales(j, i - 2)
        Next j
        sales_tax_sum = sales_tax_sum + Abs(suma)
    End If
Next i

For i = 38 To 39
suma = 0
For j = LBound(sales, 1) To UBound(sales, 1)
    suma = suma + sales(j, i - 2)
Next j
    sales_tax_sum = sales_tax_sum - Abs(suma)
Next i


purch_tax_sum = 0
For i = 42 To 50
    suma = 0
    If i = 44 Or i = 46 Or i = 47 Or i = 48 Or i = 49 Or i = 50 Then
        For j = LBound(purcharses, 1) To UBound(purcharses, 1)
            suma = suma + purcharses(j, i - 35)
        Next j
        purch_tax_sum = purch_tax_sum + Abs(suma)
    End If
Next i


' Write results into JPK sheet:
Worksheets(jpk_sheet).Range(Worksheets(jpk_sheet).Cells(jpk_row, column_numer), _
    Worksheets(jpk_sheet).Cells(jpk_row + sales_counter, column_numer + 36)) = sales
'
column_numer = column_numer + 37
jpk_row = jpk_row + sales_counter - 1
Worksheets(jpk_sheet).Cells(jpk_row, column_numer).Value = sales_counter - 1

column_numer = column_numer + 1
Worksheets(jpk_sheet).Cells(jpk_row, column_numer).Value = sales_tax_sum

column_numer = column_numer + 1
jpk_row = jpk_row + 1
Worksheets(jpk_sheet).Range(Worksheets(jpk_sheet).Cells(jpk_row, column_numer), _
    Worksheets(jpk_sheet).Cells(jpk_row + purcharses_counter + 2, column_numer + 16)) = purcharses

column_numer = column_numer + 15
jpk_row = jpk_row + purcharses_counter - 1
Worksheets(jpk_sheet).Cells(jpk_row, column_numer).Value = purcharses_counter - 1

column_numer = column_numer + 1
Worksheets(jpk_sheet).Cells(jpk_row, column_numer).Value = purch_tax_sum


' Format date numbers
Worksheets(jpk_sheet).Range("r1:s1").EntireColumn.NumberFormat = "m/d/yyyy"
Worksheets(jpk_sheet).Range("t1:aw1").EntireColumn.NumberFormat = "$#,##0.00_);($#,##0.00)"
Worksheets(jpk_sheet).Range("ay1").EntireColumn.NumberFormat = "$#,##0.00_);($#,##0.00)"
Worksheets(jpk_sheet).Range("be1:bf1").EntireColumn.NumberFormat = "m/d/yyyy"
Worksheets(jpk_sheet).Range("bp1").EntireColumn.NumberFormat = "$#,##0.00_);($#,##0.00)"

' Go to JPK sheet - can be removed
Sheets(jpk_sheet).Select

' Freeze first row for convenient scroll
Worksheets(jpk_sheet).Range("A2").Select
ActiveWindow.FreezePanes = True

Worksheets(jpk_sheet).Range(Worksheets(jpk_sheet).Cells(2, 10), Worksheets(jpk_sheet).Cells(4 + purcharses_counter + sales_counter + 2, column_numer)).WrapText = False
Worksheets(pvt_sheet).Range("a1:j1").EntireColumn.WrapText = False

' Return updating to default state
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True

' Print time for benchmarking
Debug.Print "It took: " & (DateDiff("s", #1/1/1970#, Now()) - start_time) & " seconds to generate sheet with pivot"
End Sub


Function Make_sheet(name As String)
    ' Simple function to create new worksheet with given name.
    ' If worksheet with the name exists, function remove it and create new one
    ' After creation function move sheet to the last card

    If Check_worksheet(name) Then
        'Sheet exist so remove it
        Application.DisplayAlerts = False
        ActiveWorkbook.Worksheets(name).Delete
        Application.DisplayAlerts = True
        Worksheets.Add.name = name
    Else
        Worksheets.Add.name = name
    End If
     ' Push sheet to the end
     Worksheets(name).Move after:=Worksheets(Worksheets.Count)

End Function


Function Check_worksheet(name As String)
    ' Simple function to check if sheet with given name exists
    ' Returns true or false

    Dim wsTest As Worksheet
    Set wsTest = Nothing
    On Error Resume Next
    Set wsTest = ActiveWorkbook.Worksheets(name)
    On Error GoTo 0
     
    If wsTest Is Nothing Then
        Check_worksheet = False
    Else
        Check_worksheet = True
    End If
End Function


Function MonthDays(myMonth As Long) As Integer
    ' Simple function to return how many days have given month
    MonthDays = Day(DateSerial(Year(Date), myMonth + 1, 1) - 1)
End Function




