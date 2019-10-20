Attribute VB_Name = "Module1"
Sub VBAStocks()
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
    
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    
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
                Cells(j, 11).NumberFormat = "0.00%"
            End If
            Cells(j, 12) = total
            
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
        Cells(j, 11).NumberFormat = "0.00%"
    End If
    
    Cells(j, 12) = total

End Sub
