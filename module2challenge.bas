Attribute VB_Name = "Module11"
Sub summary()

Dim i As Long, j As Long

'moves through all sheets with a for next loop
Dim sh As Long
Dim sheetcount As Long
sheetcount = Sheets.Count

For sh = 1 To sheetcount
    
    'adding headers to new summary table
    Dim titles() As String, headers As String
    headers = "Ticker, Yearly Change, Percent Change, Total Stock Volume"
    titles() = Split(headers, ", ")
    i = 9
    For j = 0 To 3
        Sheets(sh).Cells(1, i).Value = titles(j)
        i = i + 1
    Next j

    'finding how many rows in dataset
    Dim finalrow As Variant
    finalrow = Sheets(sh).Cells(Rows.Count, 1).End(xlUp).Row
    
    'startblock shows first row indiv ticker data groups
    Dim startblock As Long
    startblock = 2
    
    'ticklist allows us to determine what row to add info into summary table
    Dim ticklist As Integer
    ticklist = 1
    
    For i = 2 To finalrow
        'checks for different ticker symbol and does calculations for summary table
         If Sheets(sh).Cells(i + 1, 1).Value <> Sheets(sh).Cells(i, 1).Value Then
         
            Dim openval As Variant, finalval As Variant, totalvol As Variant
            Dim yrchange As Variant, perchange As Variant
            
            'defines value at open
            openval = (Sheets(sh).Cells(startblock, 3).Value)
            
            'defines value at close at end of year
            finalval = (Sheets(sh).Cells(i, 6).Value)
            
            'final calculation of total stock volume
            totalvol = totalvol + Sheets(sh).Cells(i, 7).Value
            
            'allows correct row placement in summary table
            ticklist = ticklist + 1
            
            'filling in summary table - ticker name
            Dim tickername As String
            tickername = Sheets(sh).Cells(i, 1).Value
            Sheets(sh).Cells(ticklist, 9).Value = tickername
            
            'filling in summary table - calculations
            yrchange = (finalval - openval)
            Sheets(sh).Cells(ticklist, 10).Value = yrchange
            perchange = (yrchange / openval)
            Sheets(sh).Cells(ticklist, 11).Value = perchange
            Sheets(sh).Cells(ticklist, 11).NumberFormat = "0.00%"
            Sheets(sh).Cells(ticklist, 12).Value = totalvol
            
            'defines where the first row is of the next ticker symbol
            startblock = i + 1
            
            'resets total stock volume calculation
            totalvol = 0
        
        'if the two rows of ticker are the same it adds stock volume to totalvol
        Else
            totalvol = totalvol + Sheets(sh).Cells(i, 7).Value
        
        End If
    Next i
    

    'secondary analysis
    Dim finalsumrow As Variant
    finalsumrow = Sheets(sh).Cells(Rows.Count, 9).End(xlUp).Row

    'Format yearly change to be green if positive and red if negative
    For i = 2 To finalsumrow
        If Sheets(sh).Cells(i, 10).Value > 0 Then
            Sheets(sh).Cells(i, 10).Interior.ColorIndex = 4
        Else
            Sheets(sh).Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i

    'filling in titles for extra analysis
    Sheets(sh).Cells(2, 15).Value = "Greatest % Increase"
    Sheets(sh).Cells(3, 15).Value = "Greatest % Decrease"
    Sheets(sh).Cells(4, 15).Value = "Greatest Total Volume"
    Sheets(sh).Cells(1, 16).Value = "Ticker"
    Sheets(sh).Cells(1, 17).Value = "Value"
    
    'format two cells to be in percent format
    Sheets(sh).Range("Q2:Q3").NumberFormat = "0.00%"

    'determining values of greatest increase/decrease and total volume
    Dim grincr As Variant
    Dim grdecr As Variant
    Dim grstvol As Variant
    grincr = 0
    grdecr = 0
    grstvol = 0
       
    'greatest percent increase check
    For i = 2 To finalsumrow
        If Sheets(sh).Cells(i, 11).Value > grincr Then
            grincr = Sheets(sh).Cells(i, 11).Value
        End If
    Next i
    Sheets(sh).Cells(2, 17).Value = grincr
    'ticker for greatest percent increase
    Dim countgrin As Integer
    For i = 2 To finalsumrow
        If Sheets(sh).Cells(i, 11).Value = grincr Then
            Sheets(sh).Cells(2, 16).Value = Sheets(sh).Cells(i, 9).Value
            countgrin = countgrin + 1
        End If
    Next i
    
    'greatest percent decrease check
    For i = 2 To finalsumrow
        If Sheets(sh).Cells(i, 11).Value < grdecr Then
            grdecr = Sheets(sh).Cells(i, 11).Value
        End If
    Next i
    Sheets(sh).Cells(3, 17).Value = grdecr
    'ticker for greatest percent decrease
    Dim countgrde As Integer
    For i = 2 To finalrow
        If Sheets(sh).Cells(i, 11).Value = grdecr Then
            Sheets(sh).Cells(3, 16).Value = Sheets(sh).Cells(i, 9)
            countgrde = countgrde + 1
        End If
    Next i
    
    
    'greatest total volume check
    For i = 2 To finalsumrow
        If Sheets(sh).Cells(i, 12).Value > grstvol Then
            grstvol = Sheets(sh).Cells(i, 12).Value
        End If
    Next i
    Sheets(sh).Cells(4, 17).Value = grstvol
    'ticker for greatest total volume check
    Dim countgrstvol As Integer
    For i = 2 To finalsumrow
        If Sheets(sh).Cells(i, 12).Value = grstvol Then
            Sheets(sh).Cells(4, 16).Value = Sheets(sh).Cells(i, 9)
            countgrstvol = countgrstvol + 1
        End If
    Next i
    
    'confirming single value for table
    If countgrin > 1 Then
        MsgBox ("More than 1 stock share the greatest percent increase value in year ") & Sheets(sh).Name
    End If

    If countgrde > 1 Then
        MsgBox ("More than 1 stock share the greatest percent decrease value in year ") & Sheets(sh).Name
    End If

    If countgrstvol > 1 Then
        MsgBox ("More than 1 stock share the greatest total stock volume in year ") & Sheets(sh).Name
    End If
    countgrin = 0
    countgrde = 0
    countgrstvol = 0
    Sheets(sh).Range("A:Q").Columns.AutoFit
    
'moves to next workbook sheet
Next sh


End Sub
