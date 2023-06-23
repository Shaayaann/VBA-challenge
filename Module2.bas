Attribute VB_Name = "Module1"




Sub Stock_Market()



For Each ws In Worksheets




  Dim Ticker_Name As String
  

 
  Dim Volume_Total As Double
  Volume_Total = 0
  
  Dim Yearly_Change As Double
  
  Dim Counter As Integer
  Counter = 0
  Dim Counter2 As Integer
  Counter2 = 0
  
  Dim Percent_Change As Double
  
  Dim Greatest_Increase As Double
  Greatest_Increase = 0
  Dim Greatest_Decrease As Double
  Greatest_Decrease = 0
  Dim Greatest_Volume As Double
  Greatest_Volume = 0
  Dim LastRow As Double
  
  
  
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2


  For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


      Ticker_Name = ws.Cells(i, 1).Value

      Volume_Total = Volume_Total + ws.Cells(i, 7).Value
      
      Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(i - Counter, 3)
      
      Percent_Change = ((ws.Cells(i, 6).Value - ws.Cells(i - Counter, 3)) / ws.Cells(i - Counter, 3)) * 100

  
      ws.Range("M" & Summary_Table_Row).Value = Ticker_Name

      
      ws.Range("P" & Summary_Table_Row).Value = Volume_Total
      
      ws.Range("N" & Summary_Table_Row).Value = Yearly_Change
      
      ws.Range("O" & Summary_Table_Row).Value = Percent_Change
      
    

  


      Summary_Table_Row = Summary_Table_Row + 1
      

      Volume_Total = 0
      Counter = 0
      Counter2 = Counter2 + 1


    Else

      Volume_Total = Volume_Total + ws.Cells(i, 7).Value
      
      Counter = Counter + 1

    End If
    
    
    

  
  Next i
  
  
  
    For i = 2 To Counter2
    

      If ws.Cells(i, 15).Value > Greatest_Increase Then
      
      Greatest_Increase = ws.Cells(i, 15).Value
      ws.Cells(2, 22).Value = Greatest_Increase
      ws.Cells(2, 21).Value = ws.Cells(i, 13).Value
      
      End If
      
      
          
      If ws.Cells(i, 15).Value < Greatest_Decrease Then
      
      Greatest_Decrease = ws.Cells(i, 15).Value
      ws.Cells(3, 22).Value = Greatest_Decrease
      ws.Cells(3, 21).Value = ws.Cells(i, 13).Value
      
      End If
      
      
      If ws.Cells(i, 16).Value > Greatest_Volume Then
      
      Greatest_Volume = ws.Cells(i, 16).Value
      ws.Cells(4, 22).Value = Greatest_Volume
      ws.Cells(4, 21).Value = ws.Cells(i, 13).Value
      
      End If
       
    
    
    
    
    
    
    
    
    
    
    If ws.Cells(i, 14).Value > 0 Then
      
        ws.Cells(i, 14).Interior.ColorIndex = 4
      
    Else
      
        ws.Cells(i, 14).Interior.ColorIndex = 3
      
    End If
    
    If ws.Cells(i, 15).Value > 0 Then
      
        ws.Cells(i, 15).Interior.ColorIndex = 4
      
    Else
      
        ws.Cells(i, 15).Interior.ColorIndex = 3
      
    End If


Next i




ws.Cells(1, 13).Value = "Ticker"
ws.Cells(1, 14).Value = "Yearly Change"
ws.Cells(1, 15).Value = "Percent Change"
ws.Cells(1, 16).Value = "Total Stock Volume"
ws.Cells(1, 21).Value = "Ticker"
ws.Cells(1, 22).Value = "Value"
ws.Cells(2, 20).Value = "Greates % Increase"
ws.Cells(3, 20).Value = "Greates % Decrease"
ws.Cells(4, 20).Value = "Greates Total Volume"



Next ws

End Sub





