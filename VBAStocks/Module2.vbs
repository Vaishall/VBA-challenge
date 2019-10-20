Attribute VB_Name = "Module2"
Sub Challenge1()
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
    ticker = Cells(2, 1).Value
    YearOpen = Cells(2, 3).Value
    total = Cells(2, 7).Value
    
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
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
    Cells(2, 15) = "Greatest % Increase"
    Cells(3, 15) = "Greatest % Decrease"
    Cells(4, 15) = "Greatest Total Volume"
    
    
    
    
    'Check if we're dealing with a new stock
    Do
        current = Cells(i, 1).Value
        If Not current = ticker Then
            
            'swap to new stock
            Cells(j, 9) = ticker
            YearClose = Cells(i - 1, 6).Value
            Cells(j, 10) = YearClose - YearOpen
            
            'deal with case where stock never opens
            If YearOpen = 0 Then
                If YearClose = 0 Then
                    Cells(j, 11) = 0
                Else
                    Cells(j, 11) = Infinity
                End If
            Else
                Cells(j, 11) = (YearClose - YearOpen) / YearOpen
                If Cells(j, 11).Value > 0 Then
                    Cells(j, 11).Interior.ColorIndex = 4
                ElseIf Cells(j, 11).Value < 0 Then
                    Cells(j, 11).Interior.ColorIndex = 3
                End If
                
                
                'Challenge 1
                If Cells(j, 11).Value > bigperinc Then
                    bigperinc = Cells(j, 11).Value
                    bigincticker = ticker
                End If
                If Cells(j, 11).Value < bigperdec Then
                    bigperdec = Cells(j, 11).Value
                    bigdecticker = ticker
                End If
                
                
                Cells(j, 11).NumberFormat = "0.00%"
            End If
            Cells(j, 12) = total
            
            'Challenge 1
            If total > bigtotal Then
                bigtotal = total
                bigtotalticker = ticker
            End If
            
            total = 0
            YearOpen = Cells(i, 3).Value
            ticker = current
            j = j + 1
        End If
        
        'update values
        total = total + Cells(i, 7).Value
        i = i + 1
        'take care of if stock opens during year
        If YearOpen = 0 Then
            YearOpen = Cells(i, 3).Value
        End If
    Loop While Not IsEmpty(Cells(i, 1).Value)
    
    
    
    
    Cells(j, 9) = ticker
    YearClose = Cells(i - 1, 6).Value
    Cells(j, 10) = YearClose - YearOpen
    
    If YearOpen = 0 Then
        If YearClose = 0 Then
            Cells(j, 11) = 0
        Else
            Cells(j, 11) = Infinity
        End If
    Else
        Cells(j, 11) = (YearClose - YearOpen) / YearOpen
        If Cells(j, 11).Value > 0 Then
            Cells(j, 11).Interior.ColorIndex = 4
        ElseIf Cells(j, 11).Value < 0 Then
            Cells(j, 11).Interior.ColorIndex = 3
        End If
        
        
        'Challenge 1
        If Cells(j, 11).Value > bigperinc Then
            bigperinc = Cells(j, 11).Value
            bigincticker = ticker
        End If
        If Cells(j, 11).Value < bigperdec Then
            bigperdec = Cells(j, 11).Value
            bigdecticker = ticker
        End If
        
        
        
        Cells(j, 11).NumberFormat = "0.00%"
    End If
    
    Cells(j, 12) = total
    'Challenge 1
    If total > bigtotal Then
        bigtotal = total
        bigtotalticker = ticker
    End If
    
    
    'CHALLENGE 1
    Cells(2, 16) = bigincticker
    Cells(2, 17) = bigperinc
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 16) = bigdecticker
    Cells(3, 17) = bigperdec
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 16) = bigtotalticker
    Cells(4, 17) = bigtotal
    
    
End Sub

