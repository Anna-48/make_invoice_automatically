Attribute VB_Name = "makeInvoice"

' INVOICE TO MERCHANTS WITH TRANSACTION + MONTHLY FEE
Sub merchant_with_transaction_and_monthly_fee()
    
    'CHOOSE FOLDER TO SAVE FILES
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Dim invTempSheet As Worksheet
    Set invTempSheet = ThisWorkbook.Sheets("INV TEMPLATE")
    Dim invListSheet As Worksheet
    Set invListSheet = ThisWorkbook.Sheets("INV list")
    
    Dim lastMerchant As Integer
    lastMerchant = invListSheet.Range("A" & invListSheet.Rows.Count).End(xlUp).Row

    Dim i As Integer
    For i = 3 To lastMerchant Step 1
        
    'SET MERCHANT
        invTempSheet.Activate
        invTempSheet.Range("B13").Value = invListSheet.Range("B" & i).Value
        
    'DUPLICATE SHEET
        Dim lastSheet As Worksheet
        Set lastSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        ActiveSheet.Copy After:=lastSheet
        
        Dim newSheet As Worksheet
        Set newSheet = ActiveSheet
        newSheet.Name = invListSheet.Range("A" & i).Value
   
    'CONVERT FORMULAS TO VALUES
        newSheet.Range("B9").Value = newSheet.Range("B9").Value
        newSheet.Range("F12").Value = newSheet.Range("F12").Value
        newSheet.Range("F15").Value = newSheet.Range("F15").Value
        newSheet.Range("B19:D20").Value = newSheet.Range("B19:D20").Value
        newSheet.Range("B32").Value = newSheet.Range("B32").Value
        newSheet.Range("B34:G44").Copy
        newSheet.Range("B34:G44").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        newSheet.Range("A30").Select
        
    'DELETE BLANKS
        Dim totalCell As Range
        Set totalCell = newSheet.Range("C" & newSheet.Rows.Count).End(xlUp)
        
        Dim totalRow As Integer
        totalRow = totalCell.Row
        
        Dim lastRow As Integer
        lastRow = totalCell.End(xlUp).Row
        
        Dim startRow As Integer
        startRow = lastRow + 2
        
        Dim endRow As Integer
        endRow = totalRow - 1
        
        If startRow < endRow Then
            newSheet.Rows(startRow & ":" & endRow).Delete
        End If

    'PRINT
        Dim file_Name As String
        file_Name = invListSheet.Range("A" & i).Value & ".pdf"
        newSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                     Filename:=folderPath & "\" & file_Name, _
                                     Quality:=xlQualityStandard, _
                                     IncludeDocProperties:=True, _
                                     IgnorePrintAreas:=False, _
                                     OpenAfterPublish:=False
    Next i
End Sub


' INVOICE TO MERCHANTS WITH JUST MONTHLY FEE
Sub merchant_with_just_monthly_fee()
    
    'CHOOSE FOLDER TO SAVE FILES
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Dim invTempSheet As Worksheet
    Set invTempSheet = ThisWorkbook.Sheets("INV TEMPLATE")
    Dim invListSheet As Worksheet
    Set invListSheet = ThisWorkbook.Sheets("INV list")
    
    Dim lastMerchant As Integer
    lastMerchant = invListSheet.Range("D" & invListSheet.Rows.Count).End(xlUp).Row

    Dim i As Integer
    For i = 3 To lastMerchant
    
    'SET MERCHANT
        invTempSheet.Activate
        invTempSheet.Range("B13").Value = invListSheet.Range("E" & i).Value
       
    'DUPLICATE SHEET
        Dim lastSheet As Worksheet
        Set lastSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        
        invTempSheet.Copy After:=lastSheet
        
        Dim newSheet As Worksheet
        Set newSheet = ActiveSheet
        newSheet.Name = invListSheet.Range("D" & i).Value
   
    'CONVERT FORMULAS TO VALUES
        newSheet.Range("B9").Value = newSheet.Range("B9").Value
        newSheet.Range("F12").Value = newSheet.Range("F12").Value
        newSheet.Range("F15").Value = newSheet.Range("F15").Value
        newSheet.Range("B19:D20").Value = newSheet.Range("B19:D20").Value
        
    'PRINT
        Dim file_Name As String
        file_Name = invListSheet.Range("D" & i).Value & ".pdf"
        newSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                     Filename:=folderPath & "\" & file_Name, _
                                     Quality:=xlQualityStandard, _
                                     IncludeDocProperties:=True, _
                                     IgnorePrintAreas:=False, _
                                     OpenAfterPublish:=False, _
                                     From:=1, To:=1
    Next i
End Sub
