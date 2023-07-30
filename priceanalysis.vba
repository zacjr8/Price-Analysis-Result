Sub MoveSheetsBeforeFinalAnalysis()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim firstCell As Range
    Dim lastCell As Range
    Dim dataRange As Range
    Dim finalAnalysisSheet As Worksheet
    Dim countBeforeFinalAnalysis As Integer
    Dim result As Variant
    
    ' Set the workbook and final analysis sheet references
    Set wb = ThisWorkbook
    Set finalAnalysisSheet = wb.Sheets("Final Analysis")
    
    'Delete final analayis main if present
    On Error Resume Next ' Ignore error if worksheet doesn't exist
    Application.DisplayAlerts = False ' Turn off the delete confirmation prompt
    ThisWorkbook.Sheets("Final Analysis Main").Delete
    Application.DisplayAlerts = True ' Turn on the delete confirmation prompt
    On Error GoTo 0 ' Reset error handling

    
    ' Loop through all the sheets in reverse order
    For i = 1 To wb.Sheets.Count
        Set ws = wb.Sheets(i)
        ' Check if the sheet is new and comes after the final analysis sheet
        If ws.Name <> "Final Analysis" And ws.Index > finalAnalysisSheet.Index Then
            ' Move the new sheet before the final analysis sheet
            ws.Move Before:=finalAnalysisSheet
        End If
    Next i
    
    ' Get the count of sheets before the final analysis sheet
    countBeforeFinalAnalysis = finalAnalysisSheet.Index - 1
    
    ' Display the count in a message box
    'MsgBox "Number of sheets before Final Analysis: " & countBeforeFinalAnalysis
    
    Set NewSheet = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    NewSheet.Name = "Final Analysis Main"
    
    
    ' Set the workbook and worksheet references
    Set wb = ThisWorkbook
    Set ws = wb.Sheets(1) ' Assuming the first sheet is the source sheet
    
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    
    Set sourceSheet = wb.Sheets("Final Analysis")
    Set destinationSheet = wb.Sheets("Final Analysis Main")
    
    Dim sourceRange As Range
    sourceSheet.Range("A1:G5").Copy
    
    
    Dim destinationRange As Range
    
    destinationSheet.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    destinationSheet.Range("A1").PasteSpecial Paste:=xlPasteFormats
   
    
    'destinationSheet.Columns = AutoFit
    
    
    
    ' Find the row with "Name" in the first column
    Dim nameRow As Range
    Set nameRow = ws.Cells.Find(what:="Name", LookIn:=xlValues, lookat:=xlWhole)
    
    If Not nameRow Is Nothing Then
        ' Find the last row in the first column (using xlDown)
        lastrow = ws.Cells(nameRow.Row, 1).End(xlDown).Row
        
        refRow = lastrow
        
        ' Set the destination sheet (change "Final Analysis Main" to the desired sheet name)
        Set NewSheet = wb.Sheets("Final Analysis Main")
        
        ' Copy the data to the new sheet
        ws.Range(ws.Cells(nameRow.Row, 1), ws.Cells(lastrow, 3)).Copy NewSheet.Range("A6")
        
        ' Optional: Autofit columns in the new sheet
        NewSheet.Columns.AutoFit
    Else
        MsgBox "Data not found. Make sure the first column contains 'Name'."
    End If

    
    Dim lastColumn As Long
    Dim sheet As Worksheet
    ' Find the last column in the 6th row of Final Analysis Main
    lastColumn = wb.Sheets("Final Analysis Main").Cells(7, wb.Sheets("Final Analysis Main").Columns.Count).End(xlToLeft).Column
    
    
    Dim refColumn As Long
    refColumn = lastColumn

    ' Loop through each sheet
    For Each sheet In wb.Sheets
        ' Check if the sheet is not "Final Analysis Main" and the 6th row is empty
        If sheet.Name <> "Final Analysis Main" And sheet.Name <> "Final Analysis" And wb.Sheets("Final Analysis Main").Cells(6, lastColumn + 1).Value = "" Then
            ' Copy the sheet name to the next empty cell in the 6th row of Final Analysis Main
            wb.Sheets("Final Analysis Main").Cells(6, lastColumn + 1).Value = sheet.Name
            lastColumn = lastColumn + 1 ' Move to the next column
        End If
    Next sheet
    
    wb.Sheets("Final Analysis Main").Columns(2).Delete shift:=xlToLeft
  
  
  
''''''''''''''''''''''''''''''''''''''''''''XLOOKUP EACH SHEET''''''''''''''''''''''''''''''''
  
  
    ' Loop through all sheets
    For Each ws In wb.Sheets
        

        For i = 7 To lastrow
            
            
            ' Set the lookup value to the second column of the 7th row in "Final Analysis Main" sheet
            Set lookupValue = wb.Sheets("Final Analysis Main").Cells(i, 2)
        
            ' Check if the sheet is not "Final Analysis Main" and "Final Analysis"
            If ws.Name <> "Final Analysis Main" And ws.Name <> "Final Analysis" And ws.Name <> "Final Analysis" Then
                ' Find the "Price in €" column in the current sheet (assuming it's in the first row)
                Set returnArray = ws.Rows(6).Find(what:="Price in €", LookIn:=xlValues, lookat:=xlWhole)
                
                
                
                ' Find the "Price in €" column in the first row of the specified sheet
                On Error Resume Next
                Set priceColumn = ws.Rows(6).Find(what:="Price in €", LookIn:=xlValues, lookat:=xlWhole)
                On Error GoTo 0
        
                If Not priceColumn Is Nothing Then
                    Dim lastrowP As Long
                    ' Find the last row in the "Price in €" column using xlDown
                    lastrowP = ws.Cells(ws.Rows.Count, priceColumn.Column).End(xlUp).Row
            
                    ' Get the range of the "Price in €" column from the first row to the last row
                    Dim priceRange As Long
                    priceRange = priceColumn.Column
                End If
                
                Set returnArray = ws.Range(ws.Cells(7, priceRange), ws.Cells(lastrowP, priceRange))
                
                
                ' Find the "Art.-Nr" column in the first row of the specified sheet
                On Error Resume Next
                Set artNrColumn = ws.Rows(6).Find(what:="Art.-Nr", LookIn:=xlValues, lookat:=xlWhole)
                On Error GoTo 0
        
                If Not artNrColumn Is Nothing Then
            
                    ' Get the range of the "Price in €" column from the first row to the last row
                    Dim artNrRange As Long
                    artNrRange = artNrColumn.Column
                End If
                
                Set lookupArray = ws.Range(ws.Cells(7, artNrRange), ws.Cells(lastrowP, artNrRange))
                
                  ' Perform the XLookup
                    On Error Resume Next ' Ignore error if lookupValue is not found
                    result = Application.WorksheetFunction.XLookup(lookupValue, lookupArray, returnArray)
                    On Error GoTo 0 ' Restore normal error handling
                    ' Check if the XLookup result is found or not
                    If Not IsError(result) Then
                        ' XLookup was successful, and result contains the matched value
                        wb.Sheets("Final Analysis Main").Cells(i, refColumn).Value = result ' Write the result in the adjacent cell
                        wb.Sheets("Final Analysis Main").Cells(i, refColumn).Font.Bold = True
                    Else
                        ' XLookup didn't find a match, handle this case accordingly
                        wb.Sheets("Final Analysis Main").Cells(i, refColumn).Value = "Not Found" ' Write a default value or handle as needed
                    End If
                
                
                
                End If
            Next i
            refColumn = refColumn + 1
    Next ws
        
''''''''''''''''''''''''''''''''''''XLOOKUP FOR SUPPLIER'''''''''''''''''''''''''''''''''''''''''''''''

    Set ws = ThisWorkbook.Sheets("Final Analysis Main")
    Dim colToInsert As Long
    colToInsert = 2
    ws.Columns(colToInsert + 1).Insert shift:=xlToRight
    ws.Cells(6, colToInsert + 1).Value = "Supplier"
    Set ws = wb.Sheets(1) ' Assuming the first sheet is the source sheet
    Set wb = ThisWorkbook
    ' Loop through all sheets
    For Each Worksheet In ThisWorkbook.Sheets
        For i = 7 To lastrow
            
            
            ' Set the lookup value to the second column of the 7th row in "Final Analysis Main" sheet
            Set lookupValue = wb.Sheets("Final Analysis Main").Cells(i, 2)
        
            ' Check if the sheet is not "Final Analysis Main" and "Final Analysis"
            If ws.Name <> "Final Analysis Main" And ws.Name <> "Final Analysis" Then
                ' Find the "Supplier" column in the current sheet (assuming it's in the first row)
                Set returnArray = ws.Rows(6).Find(what:="Supplier", LookIn:=xlValues, lookat:=xlWhole)
                
                
                
                ' Find the "Price in €" column in the first row of the specified sheet
                On Error Resume Next
                Set priceColumn = ws.Rows(6).Find(what:="Supplier", LookIn:=xlValues, lookat:=xlWhole)
                On Error GoTo 0
        
                If Not priceColumn Is Nothing Then
                    ' Find the last row in the "Supplier" column using xlDown
                    lastrowP = ws.Cells(ws.Rows.Count, priceColumn.Column).End(xlUp).Row
            
                    ' Get the range of the "Supplier" column from the first row to the last row
                    priceRange = priceColumn.Column
                End If
                
                Set returnArray = ws.Range(ws.Cells(7, priceRange), ws.Cells(lastrowP, priceRange))
                
                
                ' Find the "Art.-Nr" column in the first row of the specified sheet
                On Error Resume Next
                Set artNrColumn = ws.Rows(6).Find(what:="Art.-Nr", LookIn:=xlValues, lookat:=xlWhole)
                On Error GoTo 0
        
                If Not artNrColumn Is Nothing Then
            
                    ' Get the range of the "Price in €" column from the first row to the last row
                    artNrRange = artNrColumn.Column
                End If
                
                Set lookupArray = ws.Range(ws.Cells(7, artNrRange), ws.Cells(lastrowP, artNrRange))
                
                  ' Perform the XLookup
                    On Error Resume Next ' Ignore error if lookupValue is not found
                    result = Application.WorksheetFunction.XLookup(lookupValue, lookupArray, returnArray)
                    On Error GoTo 0 ' Restore normal error handling
                    ' Check if the XLookup result is found or not
                    If Not IsError(result) Then
                        ' XLookup was successful, and result contains the matched value
                        wb.Sheets("Final Analysis Main").Cells(i, 3).Value = result ' Write the result in the adjacent cell
                    Else
                        ' XLookup didn't find a match, handle this case accordingly
                        wb.Sheets("Final Analysis Main").Cells(i, 3).Value = "Not Found" ' Write a default value or handle as needed
                    End If
                
                
                
                End If
            Next i
            refColumn = refColumn + 1
        Next Worksheet
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        ' Find the last column in the 6th row (assuming it contains the data)
        lastColumn = wb.Sheets("Final Analysis Main").Cells(6, ws.Columns.Count).End(xlToLeft).Column
        
        ' Set the source and destination ranges
        Dim sourceRanges As Range
        Dim destinationRanged As Range
        Set sourceRanges = wb.Sheets("Final Analysis Main").Range(wb.Sheets("Final Analysis Main").Cells(6, 1), wb.Sheets("Final Analysis Main").Cells(lastrow, 3))
        Set destinationRanged = wb.Sheets("Final Analysis Main").Range(wb.Sheets("Final Analysis Main").Cells(6, 4), wb.Sheets("Final Analysis Main").Cells(lastrow, lastColumn))
    
        ' Copy the column format from sourceRange to destinationRange
        For i = 1 To destinationRanged.Columns.Count
            sourceRanges.Columns(i).Copy
            destinationRanged.Columns(i).PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
        Next i
    
        ' Wrap text in each column according to the column header in row 6
        For i = 1 To destinationRanged.Columns.Count
            destinationRanged.Columns(i).WrapText = wb.Sheets("Final Analysis Main").Cells(6, i + 3).WrapText
            destinationRanged.Columns(i).EntireColumn.AutoFit
            destinationRanged.Columns(i).Font.Bold = False
        Next i
        
        
        
        
        
''''''''''''''''''''''''''''''''''''Total of each column''''''''''''''''''''''''''''''''''''''''''''''''''''''#

        wb.Sheets("Final Analysis Main").Cells(lastrow + 1, 1).Value = "Total Material Cost"
        wb.Sheets("Final Analysis Main").Cells(lastrow + 1, 1).Font.Bold = True
        For i = 4 To lastColumn
            'Calculate the sum of the column
            Dim sumResult As Double
            sumResult = Application.WorksheetFunction.Sum(wb.Sheets("Final Analysis Main").Range(wb.Sheets("Final Analysis Main").Cells(7, i), wb.Sheets("Final Analysis Main").Cells(lastrow, i)))
            
            'Place the sum result in the cell at lastrow+1
            wb.Sheets("Final Analysis Main").Cells(lastrow + 1, i).Value = sumResult
            wb.Sheets("Final Analysis Main").Cells(lastrow + 1, i).Font.Bold = True
        Next i
        
        ' Loop through each row from row 7 to lastRow

        Dim minPrice As Double
        Dim minPriceColumn As Long
        Dim j As Long
        lastColumn = wb.Sheets("Final Analysis Main").Cells(7, wb.Sheets("Final Analysis Main").Columns.Count).End(xlToLeft).Column
        For i = 7 To lastrow + 1
            ' Initialize the minimum price and the column number of the minimum price
            minPrice = wb.Sheets("Final Analysis Main").Cells(i, 4).Value ' Assuming the first price is in column D (4th column)
            minPriceColumn = 4
            
            ' Find the minimum price and its corresponding column number in the current row
            For j = 5 To lastColumn ' Start from the 5th column (assuming the prices start from column E)
                If wb.Sheets("Final Analysis Main").Cells(i, j).Value < minPrice Then
                    minPrice = wb.Sheets("Final Analysis Main").Cells(i, j).Value
                    minPriceColumn = j
                End If
            Next j
            
            ' Highlight the cell with the minimum price using light green color
            wb.Sheets("Final Analysis Main").Cells(i, minPriceColumn).Interior.Color = RGB(198, 239, 206)
        Next i
        
        
''''''''''''''''''''''''''''''''''Other Cost Comparison''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        wb.Sheets("Final Analysis").Range("A25:A32").Copy wb.Sheets("Final Analysis Main").Range("A" & lastrow + 4)
        Dim lastRowLast As Long
        Dim lastrowCop As Long
        lastRowLast = wb.Sheets("Final Analysis Main").Cells(lastrow + 4, 1).End(xlDown).Row
        lastrowCop = lastRowLast - lastrow
        
        Application.CutCopyMode = False
        
        wb.Sheets("Final Analysis Main").Columns(1).EntireColumn.AutoFit
        
        Dim lastColumnLast As Long
        lastColumnLast = wb.Sheets("Final Analysis Main").Cells(lastrow + 4, wb.Sheets("Final Analysis Main").Columns.Count).End(xlToLeft).Column
        
        ' Loop through each sheet
        For Each sheet In wb.Sheets
            ' Check if the sheet is not "Final Analysis Main" and the 6th row is empty
            If sheet.Name <> "Final Analysis Main" And sheet.Name <> "Final Analysis" And wb.Sheets("Final Analysis Main").Cells(lastrow + 4, lastColumnLast + 1).Value = "" Then
                ' Copy the sheet name to the next empty cell in the 6th row of Final Analysis Main
                wb.Sheets("Final Analysis Main").Cells(lastrow + 5, lastColumnLast + 1).Value = sheet.Name
                lastColumnLast = lastColumnLast + 1 ' Move to the next column
            End If
        Next sheet
        
        For i = 1 To lastColumnLast
            wb.Sheets("Final Analysis Main").Columns(i).EntireColumn.AutoFit
        Next i
        
        wb.Sheets("Final Analysis Main").Range(wb.Sheets("Final Analysis Main").Cells(6, 1), wb.Sheets("Final Analysis Main").Cells(lastrowCop + 3, destinationRanged.Columns.Count - 1)).Copy
        wb.Sheets("Final Analysis Main").Range(wb.Sheets("Final Analysis Main").Cells(lastrow + 5, 1), wb.Sheets("Final Analysis Main").Cells(lastRowLast, lastColumnLast)).PasteSpecial Paste:=xlPasteFormats
        
        

'''''''''''''''''''''''''''''''''''XLOOKUP for Other Cost Comparison'''''''''''''''''''''''''''''''''''''''''

         ' Loop through all sheets
        Dim k As Integer
        k = 2
        For Each Worksheet In ThisWorkbook.Sheets
            If Worksheet.Name <> "Final Analysis Main" And Worksheet.Name <> "Final Analysis" Then
              Dim lastrowUp As Long
              lastrowUp = ThisWorkbook.Sheets("Final Analysis Main").Cells(lastrow + 100, 1).End(xlUp).Row
              For i = lastrow + 6 To lastrowUp
                  
                  
                  ' Set the lookup value to the second column of the 7th row in "Final Analysis Main" sheet
                  Set lookupValue = ThisWorkbook.Sheets("Final Analysis Main").Cells(i, 1)
              
                  ' Check if the sheet is not "Final Analysis Main" and "Final Analysis"
                      
                      Set returnArray = Worksheet.Columns(3)
                      
                      
                      
                      Set lookupArray = Worksheet.Columns(1)
                      
                        ' Perform the XLookup
                          On Error Resume Next ' Ignore error if lookupValue is not found
                          result = Application.WorksheetFunction.XLookup(lookupValue, lookupArray, returnArray)
                          On Error GoTo 0 ' Restore normal error handling
                          ' Check if the XLookup result is found or not
                          If Not IsError(result) Then
                              ' XLookup was successful, and result contains the matched value
                              ThisWorkbook.Sheets("Final Analysis Main").Cells(i, k).Value = result ' Write the result in the adjacent cell
                              ThisWorkbook.Sheets("Final Analysis Main").Cells(i, k).Font.Bold = True
                          Else
                              ' XLookup didn't find a match, handle this case accordingly
                              ThisWorkbook.Sheets("Final Analysis Main").Cells(i, k).Value = "Not Found" ' Write a default value or handle as needed
                          End If
    
              Next i
              k = k + 1
            End If
    Next Worksheet
    
    
    For i = lastrow + 6 To lastrowUp
            ' Initialize the minimum price and the column number of the minimum price
            minPrice = wb.Sheets("Final Analysis Main").Cells(i, 2).Value ' Assuming the first price is in column D (4th column)
            minPriceColumn = 2
            
            ' Find the minimum price and its corresponding column number in the current row
            For j = 3 To lastColumnLast ' Start from the 5th column (assuming the prices start from column E)
                If wb.Sheets("Final Analysis Main").Cells(i, j).Value < minPrice Then
                    minPrice = wb.Sheets("Final Analysis Main").Cells(i, j).Value
                    minPriceColumn = j
                End If
            Next j
            
            ' Highlight the cell with the minimum price using light green color
            wb.Sheets("Final Analysis Main").Cells(i, minPriceColumn).Interior.Color = RGB(198, 239, 206)
    Next i
    
    ThisWorkbook.Sheets("Final Analysis Main").Rows(lastrowUp - 1).Delete
    ThisWorkbook.Sheets("Final Analysis Main").Range("A2").Select
    

    
End Sub

