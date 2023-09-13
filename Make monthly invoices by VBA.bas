Attribute VB_Name = "makeInvoice"
'Invoicing merchants having transactions during the month
Sub invoiceMerHavingTrx()
    
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
    For i = 3 To lastMerchant
    
    'Except 34000313 (Merchant with RR)
        If invListSheet.Range("A" & i).Value = 34000313 Then
            ThisWorkbook.Sheets("RR TEMPLATE").Activate
        Else
        
        'SET MERCHANT
            invTempSheet.Activate
            invTempSheet.Range("A12").Value = invListSheet.Range("B" & i).Value
        End If
        
        'DUPLICATE SHEET
        Dim lastSheet As Worksheet
        Set lastSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        ActiveSheet.Copy After:=lastSheet
        
        Dim newSheet As Worksheet
        Set newSheet = ActiveSheet
        newSheet.Name = invListSheet.Range("A" & i).Value
   
    'CONVERT FORMULAS TO VALUES
        newSheet.Range("H7:H8").Value = newSheet.Range("H7:H8").Value
        newSheet.Range("C16:E17").Value = newSheet.Range("C16:E17").Value
        newSheet.Range("A44").Value = newSheet.Range("A44").Value
        newSheet.Range("A47:H77").Copy
        newSheet.Range("A47:H77").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        newSheet.Range("A42").Select
        
    'DELETE BLANKS
        Dim startRow As Integer
        startRow = newSheet.Range("B80").End(xlUp).Row + 2
        
        If startRow <= 78 Then
            newSheet.Rows(startRow & ":" & 78).Delete
        End If

    'PRINT
        Dim file_Name As String
        file_Name = invListSheet.Range("C" & i).Value & ".pdf"
        newSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                     Filename:=folderPath & "\" & file_Name, _
                                     Quality:=xlQualityStandard, _
                                     IncludeDocProperties:=True, _
                                     IgnorePrintAreas:=False, _
                                     OpenAfterPublish:=False
    Next i
End Sub

------------------------------------

'Invoicing merchants having no transaction but monthly subcription fee
Sub invoiceMerHavingSubFee()
    
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
    lastMerchant = invListSheet.Range("E" & invListSheet.Rows.Count).End(xlUp).Row

    Dim i As Integer
    For i = 3 To lastMerchant
    
    'SET MERCHANT
        invTempSheet.Activate
        invTempSheet.Range("A12").Value = invListSheet.Range("F" & i).Value
       
    'DUPLICATE SHEET
        Dim lastSheet As Worksheet
        Set lastSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        
        invTempSheet.Copy After:=lastSheet
        
        Dim newSheet As Worksheet
        Set newSheet = ActiveSheet
        newSheet.Name = invListSheet.Range("E" & i).Value
   
    'CONVERT FORMULAS TO VALUES
        newSheet.Range("H7:H8").Value = newSheet.Range("H7:H8").Value
        newSheet.Range("C16:E17").Value = newSheet.Range("C16:E17").Value

    'PRINT
        Dim file_Name As String
        file_Name = invListSheet.Range("G" & i).Value & ".pdf"
        newSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                     Filename:=folderPath & "\" & file_Name, _
                                     Quality:=xlQualityStandard, _
                                     IncludeDocProperties:=True, _
                                     IgnorePrintAreas:=False, _
                                     OpenAfterPublish:=False, _
                                     From:=1, To:=1
    Next i
End Sub

------------------------------------

'Invoicing 34000313
Sub invoiceMerWithRR()
    
    'CHOOSE FOLDER TO SAVE FILES
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    'SELECT SHEET
    Dim RRSheet As Worksheet
    Set RRSheet = ThisWorkbook.Sheets("34000313")
   
    'CONVERT FORMULAS TO VALUES
        RRSheet.Range("H7:H8").Value = RRSheet.Range("H7:H8").Value
        RRSheet.Range("C16:E17").Value = RRSheet.Range("C16:E17").Value
        RRSheet.Range("A44").Value = RRSheet.Range("A44").Value
        RRSheet.Range("A47:H77").Copy
        RRSheet.Range("A47:H77").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        RRSheet.Range("A42").Select
        
    'DELETE BLANKS
        
        Dim startRow As Integer
        startRow = RRSheet.Range("B80").End(xlUp).Row + 2
        
        If startRow <= 78 Then
            RRSheet.Rows(startRow & ":" & 78).Delete
        End If
        
    'PRINT
        Dim file_Name As String
        file_Name = RRSheet.Name
        RRSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                    Filename:=folderPath & "\" & file_Name, _
                                    Quality:=xlQualityStandard, _
                                    IncludeDocProperties:=True, _
                                    IgnorePrintAreas:=False, _
                                    OpenAfterPublish:=False
    
End Sub

