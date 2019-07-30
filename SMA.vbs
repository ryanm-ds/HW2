VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub SA():

    Dim Ticker As String
    Dim TSV As Double
   
    Dim SRT As Integer
    
    For Each ws In Worksheets
    ws.Activate
    
    TSV = 0
    SRT = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"
    
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRow
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = Cells(i, 1).Value
                TSV = TSV + Cells(i, "G").Value
                ws.Range("I" & SRT).Value = Ticker
                ws.Range("J" & SRT).Value = TSV
                SRT = SRT + 1
                TSV = 0
            Else
                TSV = TSV + Cells(i, "G").Value
            End If
            
            Next i
        Next ws
        
End Sub
