Attribute VB_Name = "ImportAndTransfer"
'Hide from User
Option Private Module

Public mbCancel As Boolean
Public ContainerID As String
Public RouteName As String
Public DateTime As String
Public IBYardSlot As String


'Take the container export that is pasted in the STS Export sheet and populate the Menu headers as well as move pertinent columns onto Data sheet
Sub ImportAndTransfer()
    
    Dim Menu, STS, Data As Worksheet
    Dim PasteRange, ExportRangeEnd, ASN As Range
    Dim DateRange, TimeRange, SortRange As Range
    Dim ExportNumRows, count As Integer
    
    Set Menu = Sheets("Menu")
    
    'Ask user to double check pasted information
    response = MsgBox("Please verify container ID with STS before printing.", 4353, "Stop and Verify")
    
    'if the user pressed "Cancel" then exit the subroutine.
    If response = 2 Then
        Call Clear.ClearAll
        Exit Sub
    End If
    
    With Menu
        
        GeneratorStart.Show
                
        ' This IF statement will catch the event thrown when user closes out of Z512Importer window with the X button
        If mbCancel = True Then
            Exit Sub
        End If
        
        'Retreive the container ID from the STS export - every entry in a single export will contain the same container ID
        'We do this first so the user can verify container ID shown with their paperwork
        ContainerID = Sheets("STS Export").Range("A2").Value
        
        'Populate container information on main screen / populate container ID and username on manifest
        .Range("B3") = ContainerID
        Worksheets("Manifest").Range("B1") = ContainerID
        .Range("D8") = RouteName
        .Range("D13") = Application.UserName
        Worksheets("Manifest").Range("B4") = Application.UserName   'Added 6/10/19 JW
        .Range("B8") = IBYardSlot
        .Range("B13") = Now
        
    End With
    
    
    Set STS = Sheets("STS Export")

    With STS

        'get number of rows in STS export
        Set ExportRangeEnd = .Range("A" & .rows.count).End(xlUp)
        ExportNumRows = ExportRangeEnd.Row

        'Transfer only the needed data columns to Data worksheet for further processing
        '***In order to re-write as little code as necessary for this change-over, the columns are being pasted in the order
        'that will accomodate the code currently in place to process bills.****
        
        Set ASN = .Range("A2:A" & ExportNumRows)
        Sheets("Data").Range("A2:A" & ExportNumRows) = ASN.Value

        Set ASN = .Range("C2:C" & ExportNumRows)
        Sheets("Data").Range("D2:D" & ExportNumRows) = ASN.Value

        Set ASN = .Range("E2:E" & ExportNumRows)
        Sheets("Data").Range("E2:E" & ExportNumRows) = ASN.Value

        Set ASN = .Range("G2:G" & ExportNumRows)
        Sheets("Data").Range("F2:F" & ExportNumRows) = ASN.Value

        Set ASN = .Range("I2:I" & ExportNumRows)
        Sheets("Data").Range("K2:K" & ExportNumRows) = ASN.Value

        Set ASN = .Range("J2:J" & ExportNumRows)
        Sheets("Data").Range("G2:G" & ExportNumRows) = ASN.Value

        Set ASN = .Range("K2:K" & ExportNumRows)
        Sheets("Data").Range("H2:H" & ExportNumRows) = ASN.Value

        Set ASN = .Range("L2:L" & ExportNumRows)
        Sheets("Data").Range("I2:I" & ExportNumRows) = ASN.Value

        Set ASN = .Range("M2:M" & ExportNumRows)
        Sheets("Data").Range("J2:J" & ExportNumRows) = ASN.Value

    End With
    
End Sub
