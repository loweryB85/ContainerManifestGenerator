Attribute VB_Name = "Clear"
'Hide from User
Option Private Module
'************************************************************************
'   ClearAll
'
'   This sub will clear all of the worksheets used with the current data
'   export - Menu, Manifest, Data, STS Export
'
'   Inputs - None
'   Outputs - None
'************************************************************************
Sub ClearAll()

    ClearMenu
    ClearManifest
    ClearData
    Sheets("STS Export").Cells.Clear
      
End Sub
'************************************************************************
'   ClearMenu
'
'   This sub will clear the Menu worksheet
'
'   Inputs - None
'   Outputs - None
'************************************************************************
Sub ClearMenu()

'Clear informational cells on menu sheet
Worksheets("Menu").Cells(3, 2).MergeArea.ClearContents  ' Clear Container ID
Worksheets("Menu").Cells(8, 2).MergeArea.ClearContents  ' Clear YardSlot
Worksheets("Menu").Cells(8, 4).MergeArea.ClearContents  ' Clear Route Name
Worksheets("Menu").Cells(13, 2).MergeArea.ClearContents 'Clear Import Time

End Sub

'************************************************************************
'   ClearData
'
'   This sub will clear everything from the data worksheet
'
'   Inputs - None
'   Outputs - None
'************************************************************************
Sub ClearData()

    'no point in dynamic range sizing for our purposes - max # of lines for a container is 500
    Sheets("Data").Range("A2:N500").ClearContents

End Sub


'************************************************************************
'   ClearManifest
'
'   This sub will clear the contents in the manifest worksheet.
'
'   Inputs - None
'   Outputs - None
'************************************************************************
Sub ClearManifest()

    'Clear the contents of the manifest
    Worksheets("Manifest").Range("A7:H32").ClearContents    'Changed Range from A7:I32 to accomodate re-organizing of merged cells  9/19/19 BL
    Worksheets("Manifest").Range("B1").ClearContents
    Worksheets("Manifest").Range("B2").ClearContents
    Worksheets("Manifest").Range("D5").ClearContents
    
End Sub
