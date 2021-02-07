Attribute VB_Name = "Module1"
Sub Wallstreet()

Dim Rows As Long
Dim i As Long
Dim j As Integer
Dim z As Integer
Dim Volume As Double
Dim OpPrice As Double
Dim ClPrice As Double
Dim Work As Integer
'Count the Number of Sheets
Work = Application.Sheets.Count
MsgBox (Work)

For z = 1 To Work
    Worksheets(z).Activate
    'Enter Summary Table Heading
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Counts the number of Rows and initialize Volume as Zero
    Rows = ActiveSheet.UsedRange.Rows.Count
    
    j = 2
    
    For i = 2 To Rows + 1                                               'Rows+1 is needed because otherwise, The else statement is not going to check the last Ticker prices and volumes
        
        If Cells(i, 1).Value = Cells(i - 1, 1).Value Then               'Checks If the current Cell has the same ticker as the last cell
            Volume = Volume + Cells(i, 7).Value                         'Adds Volumes
        
        ElseIf i = 2 Then                                               'This If is needed because we are comparing from previous rows but
                                                                        'it Provides problems when taking the header into account
            Cells(j, 9).Value = Cells(i, 1).Value
            Volume = Cells(i, 7)
            OpPrice = Cells(i, 3)
        
        Else
            'Ticker
            Cells(j + 1, 9).Value = Cells(i, 1).Value                   'Enters the Ticker, this is j+1 because of the i = 2 If
            
            'Volume
            Cells(j, 12).Value = Volume                                 'Enters Volume for that Ticker
            Volume = Cells(i, 7)                                        'Stores New Volume starting at the new Ticker
            
            'Yearly Change
            ClPrice = Cells(i - 1, 6)
            Cells(j, 10) = ClPrice - OpPrice                            'Enters Yearly Change
                
                'Color Fill Based on negative or positive change
                If Cells(j, 10) > 0 Then
                    Cells(j, 10).Interior.ColorIndex = 4
                Else
                    Cells(j, 10).Interior.ColorIndex = 3
                End If
            'and Percent Change
                If OpPrice = 0 Then                                     'Avoids error where we are deviding by zero
                    Cells(j, 11) = "Not Valid"
                Else
                    Cells(j, 11) = Cells(j, 10) / OpPrice
                End If
            Cells(j, 11).NumberFormat = "0.00%"
            OpPrice = Cells(i, 3)                                       'New Opening Price due to new Ticker
            
            
            j = j + 1
        End If
        
    Next i
Next z
End Sub
