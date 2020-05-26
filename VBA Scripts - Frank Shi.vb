Sub VBAchallenge()

Dim ticker As String
Dim i As Long
Dim begi As Long
Dim ed As Long
Dim rws As Long
Dim num_ticker As Integer
Dim ws As Worksheet
Dim total As Double

For Each ws In ThisWorkbook.Worksheets

    With ws

        rws = .Range("A2").End(xlDown).Row
        
        begi = 2
        ed = 0
        num_ticker = 2
        
        .Range("I" & num_ticker - 1).Value = "Ticker"
        .Range("J" & num_ticker - 1).Value = "Yearly Change"
        .Range("K" & num_ticker - 1).Value = "Percentage Change"
        .Range("L" & num_ticker - 1).Value = "Total Stock Volume"
        
        .Range("O" & num_ticker - 1).Value = "Ticker"
        .Range("P" & num_ticker - 1).Value = "Value"
        
        .Range("N" & num_ticker).Value = "Greatest % increase"
        .Range("N" & num_ticker + 1).Value = "Greatest % decrease"
        .Range("N" & num_ticker + 2).Value = "Greatest total volume"
        
        For i = 2 To rws
        
            total = total + .Range("G" & i).Value
            
            If .Cells(i, 1) <> .Cells(i + 1, 1) Then
            
                .Range("I" & num_ticker) = .Cells(i, 1).Value
                
                ed = i
                
                .Range("J" & num_ticker) = .Range("F" & ed).Value - .Range("C" & begi).Value
                
                
                    If .Range("C" & begi).Value <> 0 Then
                    
                        .Range("K" & num_ticker) = Format(Round((.Range("F" & ed).Value - .Range("C" & begi).Value) / .Range("C" & begi).Value, 5), "Percent")

                    
                    Else: .Range("K" & num_ticker) = 0
                    
                    End If
                
                .Range("L" & num_ticker) = total
                
                total = 0
                
                    If .Range("J" & num_ticker) > 0 Then
            
                        .Range("J" & num_ticker).Interior.ColorIndex = 4
                    
                    ElseIf .Range("J" & num_ticker) < 0 Then
                        
                        .Range("J" & num_ticker).Interior.ColorIndex = 3
            
                    End If
                    
                num_ticker = num_ticker + 1
                
                begi = ed + 1
               
            End If
    
        Next i
    
    End With
    
Next

End Sub


Sub VBAchallenge2()

Dim num As Integer
Dim ws As Worksheet
Dim Max_v As Double
Dim Min_v As Double
Dim Max_t As Double
Dim rws As Integer
Dim i As Integer
Dim Max_v_row As String
Dim Min_v_row As String
Dim Max_t_row As String

    For Each ws In ThisWorkbook.Worksheets
    
        Max_v_row = 0
        Min_v_row = 0
        Max_t_row = 0
            
        Max_v = 0
        Min_v = 0
        Max_t = 0
        
        With ws
        
            rws = .Range("K2").End(xlDown).Row
            
            num = 2
            
            For i = 2 To rws
            
                If .Cells(i, 11).Value >= Max_v Then
                
                        Max_v = .Cells(i, 11).Value
                        Max_v_row = .Cells(i, 9).Value
                        
                End If
                
                If .Cells(i, 11).Value <= Min_v Then
                
                        Min_v = .Cells(i, 11).Value
                        Min_v_row = .Cells(i, 9).Value
                        
                End If
                
                
                If .Cells(i, 12).Value >= Max_t Then
                
                    Max_t = .Cells(i, 12).Value
                    Max_t_row = .Cells(i, 9).Value
                    
                End If
                
            Next i
            
            .Range("O2").Value = Max_v_row
            .Range("O3").Value = Min_v_row
            .Range("O4").Value = Max_t_row
            
            .Range("P2").Value = Format(Max_v, "Percent")
            .Range("P3").Value = Format(Min_v, "Percent")
            .Range("P4").Value = Max_t
            
        End With
        
    Next

End Sub

