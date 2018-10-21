VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub homework2():
For wsn = 1 To Worksheets.Count
Set ws = Worksheets(wsn)
 ws.Range("I1").Value = ws.Range("A1").Value
 ws.Range("J1").Value = "Total Stock Volume"
 Dim category As Double
 Dim summary As Integer
 Dim samecategory As Integer
 
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Total Stock Volme"
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 summary = 2
 category = 0
 samecategory = 1

   For i = 2 To lastrow
     If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
    category = category + ws.Cells(i, 7).Value
    samecategory = samecategory + 1
  Else
    ws.Cells(summary, 9).Value = ws.Cells(i, 1).Value
   category = category + ws.Cells(i, 7).Value
   ws.Cells(summary, 10).Value = category
    category = 0
    samecategory = 1
   summary = summary + 1
   End If
 Next i

Next wsn
End Sub
 
