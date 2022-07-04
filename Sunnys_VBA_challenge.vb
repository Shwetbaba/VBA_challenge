Option Explicit

Sub Loop_Solved()


    Dim CurrentWs As Worksheet
    
    For Each CurrentWs In Worksheets
    
    
        Dim Ticker_Name As String
        Ticker_Name = " "
        

        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        

        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim Yearly_Change As Double
        Yearly_Change = 0
        Dim Percent_Change As Double
        Percent_Change = 0
        

        Dim MAX_TICKER_NAME As String
        MAX_TICKER_NAME = " "
        Dim MIN_TICKER_NAME As String
        MIN_TICKER_NAME = " "
        Dim MAX_PERCENT As Double
        MAX_PERCENT = 0
        Dim MIN_PERCENT As Double
        MIN_PERCENT = 0
        Dim MAX_VOLUME_TICKER As String
        MAX_VOLUME_TICKER = " "
        Dim MAX_VOLUME As Double
        MAX_VOLUME = 0
   
 
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        

        Dim LastRow As Long
        Dim i As Long
        
        LastRow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row


            CurrentWs.Range("I1").Value = "Ticker"
            CurrentWs.Range("J1").Value = "Yearly Change"
            CurrentWs.Range("K1").Value = "Percent Change"
            CurrentWs.Range("L1").Value = "Total Stock Volume"
            

            CurrentWs.Range("O2").Value = "Greatest % Increase"
            CurrentWs.Range("O3").Value = "Greatest % Decrease"
            CurrentWs.Range("O4").Value = "Greatest Total Volume"
            CurrentWs.Range("P1").Value = "Ticker"
            CurrentWs.Range("Q1").Value = "Value"
    
        

        Open_Price = CurrentWs.Cells(2, 3).Value
        

        For i = 2 To LastRow
        

            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
            

                Ticker_Name = CurrentWs.Cells(i, 1).Value
                

                Close_Price = CurrentWs.Cells(i, 6).Value
                Yearly_Change = Close_Price - Open_Price
                

                If Open_Price <> 0 Then
                    Percent_Change = (Yearly_Change / Open_Price) * 100
                End If
                

                Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
              
                

                CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name

                CurrentWs.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
    

                If (Yearly_Change > 0) Then
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
                ElseIf (Yearly_Change <= 0) Then
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.Color = vbRed
                    
                End If
                

                CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
  
                CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                

                Summary_Table_Row = Summary_Table_Row + 1
                
                Yearly_Change = 0
             
                Close_Price = 0
           
                Open_Price = CurrentWs.Cells(i + 1, 3).Value
              
        
                If (Percent_Change > MAX_PERCENT) Then
                    MAX_PERCENT = Percent_Change
                    MAX_TICKER_NAME = Ticker_Name
                ElseIf (Percent_Change < MIN_PERCENT) Then
                    MIN_PERCENT = Percent_Change
                    MIN_TICKER_NAME = Ticker_Name
                End If
                       
                If (Total_Ticker_Volume > MAX_VOLUME) Then
                    MAX_VOLUME = Total_Ticker_Volume
                    MAX_VOLUME_TICKER = Ticker_Name
                End If
                
          
                Percent_Change = 0
                Total_Ticker_Volume = 0
                
            
   
            Else
             
                Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
            End If
      
      
      
            
                CurrentWs.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
                CurrentWs.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
                CurrentWs.Range("P2").Value = MAX_TICKER_NAME
                CurrentWs.Range("P3").Value = MIN_TICKER_NAME
                CurrentWs.Range("Q4").Value = MAX_VOLUME
                CurrentWs.Range("P4").Value = MAX_VOLUME_TICKER
                
      
            
            CurrentWs.Range("I1:Q1").Font.Name = "Arial"
            CurrentWs.Range("I1:Q1").Font.Size = "16"
            CurrentWs.Range("I1:Q1").Font.Bold = "True"
            CurrentWs.Range("I1:Q1").EntireColumn.HorizontalAlignment = xlCenter
            CurrentWs.Range("I1:Q1").EntireColumn.AutoFit
            
        Next i


     Next CurrentWs
     
     'reference : https://freesoft.dev/program/163047389

End Sub