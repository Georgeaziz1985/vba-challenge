Attribute VB_Name = "Module1"
Sub stockanalysis()

Dim ticker As String


Dim numbertickers As Integer

Dim lastRowState As Long


Dim openingprice As Double


Dim closingprice As Double


Dim yearlychange As Double


Dim percentchange As Double


Dim totalstockvolume As Double


Dim greatestpercentincrease As Double


Dim greatestpercentincreaseticker As String


Dim greatestpercentdecrease As Double


Dim greatestpercentdecreaseticker As String


Dim greateststockvolume As Double


Dim greateststockvolumeticker As String


For Each ws In Worksheets

    
    ws.Activate


    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row

    
    ws.Range("M1").Value = "Ticker"
    ws.Range("N1").Value = "Yearly Change"
    ws.Range("O1").Value = "Percent Change"
    ws.Range("P1").Value = "Total Stock Volume"
    
    
    numbertickers = 0
    ticker = ""
    yearlychange = 0
    openingprice = 0
    percentchange = 0
    totalstock_volume = 0
    
    
    For i = 2 To lastRowState

        
        ticker = Cells(i, 1).Value
        
        
        If openingprice = 0 Then
            openingprice = Cells(i, 3).Value
        End If
        
        
        totalstockvolume = totalstockvolume + Cells(i, 7).Value
        
        
        If Cells(i + 1, 1).Value <> ticker Then
        
            numbertickers = numbertickers + 1
            Cells(numbertickers + 1, 13) = ticker
            
            
            closingprice = Cells(i, 6)
            
            
            yearlychange = closingprice - openingprice
            
            
            Cells(numbertickers + 1, 14).Value = yearlychange
            
            
            If yearlychange > 0 Then
                Cells(numbertickers + 1, 14).Interior.ColorIndex = 4
        
            ElseIf yearlychange < 0 Then
                Cells(numbertickers + 1, 14).Interior.ColorIndex = 3
            
            
            End If
            
            
            
            If openingprice = 0 Then
                percentchange = 0
            Else
                percentchange = (yearlychange / openingprice)
            End If
            
            
            
            Cells(numbertickers + 1, 15).Value = Format(percentchange, "Percent")
            
            
            
            
            
            
            openingprice = 0
            
            
            Cells(numbertickers + 1, 16).Value = totalstockvolume
            
            
            totalstockvolume = 0
        End If
        
    Next i
    
    ' Add section to display greatest percent increase, greatest percent decrease, and greatest total volume for each year.
    Range("S2").Value = "Greatest % Increase"
    Range("S3").Value = "Greatest % Decrease"
    Range("S4").Value = "Greatest Total Volume"
    Range("T1").Value = "Ticker"
    Range("U1").Value = "Value"
    
    
    lastRowState = ws.Cells(Rows.Count, "M").End(xlUp).Row
    
    
    greatestpercentincrease = Cells(2, 15).Value
    greatestpercentincreaseticker = Cells(2, 13).Value
    greatestpercentdecrease = Cells(2, 15).Value
    greatestpercentdecreaseticker = Cells(2, 13).Value
    greateststockvolume = Cells(2, 16).Value
    greateststockvolumeticker = Cells(2, 13).Value
    
    
    
    For i = 2 To lastRowState
    
        
        If Cells(i, 15).Value > greatestpercentincrease Then
            greatestpercentincrease = Cells(i, 15).Value
            greatestpercentincreaseticker = Cells(i, 13).Value
        End If
        
        
        If Cells(i, 15).Value < greatest_percent_decrease Then
            greatestpercentdecrease = Cells(i, 15).Value
            greatestpercentdecreaseticker = Cells(i, 13).Value
        End If
        
        
        If Cells(i, 16).Value > greateststockvolume Then
            greateststockvolume = Cells(i, 16).Value
            greateststockvolumeticker = Cells(i, 13).Value
        End If
        
    Next i
    
    
    Range("T2").Value = Format(greatestpercentincreaseticker, "Percent")
    Range("U2").Value = Format(greatestpercentincrease, "Percent")
    Range("T3").Value = Format(greatestpercentdecreaseticker, "Percent")
    Range("U3").Value = Format(greatestpercentdecrease, "Percent")
    Range("T4").Value = greateststockvolumeticker
    Range("U4").Value = greateststockvolume
    
Next ws
End Sub
