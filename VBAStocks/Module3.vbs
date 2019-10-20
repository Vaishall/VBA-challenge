Attribute VB_Name = "Module3"
Sub Challenge2()
    
    'Challenge 2
    Dim currentsheet As Worksheet
    
    For Each currentsheet In Worksheets
        
        
        
        
        'Define basic variables
        Dim ticker As String
        Dim current As String
        Dim YearOpen As Double
        Dim YearClose As Double
        Dim total As Double
        Dim i As Long
        Dim j As Integer
        j = 2
        i = 2
        ticker = currentsheet.Cells(2, 1).Value
        YearOpen = currentsheet.Cells(2, 3).Value
        total = currentsheet.Cells(2, 7).Value
    
        'Define Challenge Variables
        Dim bigperinc As Double
        Dim bigperdec As Double
        Dim bigtotal As Double
        bigperinc = 0
        bigperdec = 0
        bigpertotal = 0
        Dim bigincticker As String
        Dim bigdecticker As String
        Dim bigtotalticker As String
    
    
    
    
        'Set Headers and Such
        currentsheet.Cells(1, 9) = "Ticker"
        currentsheet.Cells(1, 10) = "Yearly Change"
        currentsheet.Cells(1, 11) = "Percent Change"
        currentsheet.Cells(1, 12) = "Total Stock Volume"
    
        currentsheet.Cells(1, 16) = "Ticker"
        currentsheet.Cells(1, 17) = "Value"
        currentsheet.Cells(2, 15) = "Greatest % Increase"
        currentsheet.Cells(3, 15) = "Greatest % Decrease"
        currentsheet.Cells(4, 15) = "Greatest Total Volume"
    
    
    
    
        'Check if we're dealing with a new stock
        Do
            current = currentsheet.Cells(i, 1).Value
            If Not current = ticker Then
            
                'swap to new stock
                currentsheet.Cells(j, 9) = ticker
                YearClose = currentsheet.Cells(i - 1, 6).Value
                currentsheet.Cells(j, 10) = YearClose - YearOpen
            
                'deal with case where stock never opens
                If YearOpen = 0 Then
                    If YearClose = 0 Then
                        currentsheet.Cells(j, 11) = 0
                    Else
                        currentsheet.Cells(j, 11) = Infinity
                    End If
                Else
                    currentsheet.Cells(j, 11) = (YearClose - YearOpen) / YearOpen
                    If currentsheet.Cells(j, 11).Value > 0 Then
                        currentsheet.Cells(j, 11).Interior.ColorIndex = 4
                    ElseIf currentsheet.Cells(j, 11).Value < 0 Then
                        currentsheet.Cells(j, 11).Interior.ColorIndex = 3
                    End If
                
                
                    'Challenge 1
                    If currentsheet.Cells(j, 11).Value > bigperinc Then
                        bigperinc = currentsheet.Cells(j, 11).Value
                        bigincticker = ticker
                    End If
                    If currentsheet.Cells(j, 11).Value < bigperdec Then
                        bigperdec = currentsheet.Cells(j, 11).Value
                        bigdecticker = ticker
                    End If
                
                
                    currentsheet.Cells(j, 11).NumberFormat = "0.00%"
                End If
                currentsheet.Cells(j, 12) = total
            
                'Challenge 1
                If total > bigtotal Then
                    bigtotal = total
                    bigtotalticker = ticker
                End If
                
                total = 0
                YearOpen = currentsheet.Cells(i, 3).Value
                ticker = current
                j = j + 1
            End If
        
            'update values
            total = total + currentsheet.Cells(i, 7).Value
            i = i + 1
            'take care of if stock opens during year
            If YearOpen = 0 Then
                YearOpen = currentsheet.Cells(i, 3).Value
            End If
        Loop While Not IsEmpty(currentsheet.Cells(i, 1).Value)
    
    
    
    
        currentsheet.Cells(j, 9) = ticker
        YearClose = currentsheet.Cells(i - 1, 6).Value
        currentsheet.Cells(j, 10) = YearClose - YearOpen
    
        If YearOpen = 0 Then
            If YearClose = 0 Then
                currentsheet.Cells(j, 11) = 0
            Else
                currentsheet.Cells(j, 11) = Infinity
            End If
        Else
            currentsheet.Cells(j, 11) = (YearClose - YearOpen) / YearOpen
            If currentsheet.Cells(j, 11).Value > 0 Then
                currentsheet.Cells(j, 11).Interior.ColorIndex = 4
            ElseIf currentsheet.Cells(j, 11).Value < 0 Then
                currentsheet.Cells(j, 11).Interior.ColorIndex = 3
            End If
        
        
            'Challenge 1
            If currentsheet.Cells(j, 11).Value > bigperinc Then
                bigperinc = currentsheet.Cells(j, 11).Value
                bigincticker = ticker
            End If
            If currentsheet.Cells(j, 11).Value < bigperdec Then
                bigperdec = currentsheet.Cells(j, 11).Value
                bigdecticker = ticker
            End If
        
        
        
            currentsheet.Cells(j, 11).NumberFormat = "0.00%"
        End If
    
        currentsheet.Cells(j, 12) = total
        'Challenge 1
        If total > bigtotal Then
            bigtotal = total
            bigtotalticker = ticker
        End If
    
    
        'CHALLENGE 1
        currentsheet.Cells(2, 16) = bigincticker
        currentsheet.Cells(2, 17) = bigperinc
        currentsheet.Cells(2, 17).NumberFormat = "0.00%"
        currentsheet.Cells(3, 16) = bigdecticker
        currentsheet.Cells(3, 17) = bigperdec
        currentsheet.Cells(3, 17).NumberFormat = "0.00%"
        currentsheet.Cells(4, 16) = bigtotalticker
        currentsheet.Cells(4, 17) = bigtotal
    
    
    
    Next
End Sub


