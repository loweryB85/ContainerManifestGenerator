Attribute VB_Name = "ManifestFunctions"
'Hide from User
Option Private Module


'************************************************************************
'   CopyPasteBOL
'
'   This sub will take the helper column number selected by the HCfilter
'   form and loop through the part data looking for the selected number
'   in the helper column. The corresponding rows will be pasted into the
'   Manifest sheet in rows 7-32.
'
'   Inputs - HC (integer value from HCfilter form)
'   Outputs - None
'************************************************************************
Sub CopyPasteBOL(ByVal HC As Integer)

    'Range variable to transfer data
    Dim toCopy As Range

    'variable to iterate through BOLSheet - starts at row 7 of form
    Dim ManifestRow As Integer
    ManifestRow = 7
    
    'Begin by populating the header
    PopulateHeader
    
    'Note that we only allowed for a maximum entry of 10000 lines
    For Each Cell In Sheets("Data").Range("M2:M1000")
        
        'The current row in the Data worksheet matches the BOL we are looking for
        If Cell.Value = HC Then
            
            'Assign the data from this row to a range variable
            Set toCopy = Sheets("Data").Range("D" & Cell.Row & ":K" & Cell.Row)

            'copy the contents of the range variable into the current row in the Manifest
            Worksheets("Manifest").Range("A" & ManifestRow & ":H" & ManifestRow) = toCopy.Value
                        
            'Move to the next row of the manifest
            ManifestRow = ManifestRow + 1
            
            'Added 6/5/19 - BL
            'If at end of page (after line 32) print page, clear page, start over at line 7.
            If ManifestRow = 33 Then
                
                'Print the manifest, clear it, and then reset the ManifestRow to the first line (7)
                PrintManifest
                Clear.ClearManifest
                ManifestRow = 7
                
            End If
                            
        End If
        
    Next
    
End Sub

'************************************************************************
'   CopyPasteFacilityDock
'
'   This sub will take the helper column number selected by the HCfilter
'   form and loop through the part data looking for the selected number
'   in the helper column. The corresponding rows will be pasted into the
'   Manifest sheet in rows 7-29.
'
'   Inputs - HC (integer value from HCfilter form)
'   Outputs - None
'************************************************************************
Sub CopyPasteFacilityDock(ByVal HC As Integer)

    'Range variable to transfer data
    Dim toCopy As Range

    'variable to iterate through BOLSheet - starts at row 7 of form
    Dim ManifestRow As Integer
    ManifestRow = 7
    
    'Begin by populating the header
    PopulateHeader
    
    'Note that we only allowed for a maximum entry of 10000 lines
    For Each Cell In Sheets("Data").Range("N2:N1000")
        
        'the current row in the Data sheet goes to the Dock we are looking for
        If Cell.Value = HC Then
            
            'Assign the data from this row to a range variable
            Set toCopy = Sheets("Data").Range("D" & Cell.Row & ":K" & Cell.Row)

            'copy the contents of the range variable into the current row in the Manifest
            Worksheets("Manifest").Range("A" & ManifestRow & ":H" & ManifestRow) = toCopy.Value
            
            'move to the next row of the manifest
            ManifestRow = ManifestRow + 1
                       
            'Added 6/5/19 - BL
            'If at end of page (after line 32) print page, clear page, start over at line 7.
            If ManifestRow = 33 Then
                
                'Print the manifest, clear it, and then reset the ManifestRow to the first line (7)
                PrintManifest
                Clear.ClearManifest
                ManifestRow = 7
                
            End If
            
        End If
                
    Next
    
End Sub

'************************************************************************
'   PopulateHeader
'
'   This sub will populate the user-provided information into the header
'   of the manifest worksheet, specifically container id, route name,
'   yard slot.
'
'   Inputs - None
'   Outputs - None
'************************************************************************
Sub PopulateHeader()

    Worksheets("Manifest").Range("B1") = ContainerID
    Worksheets("Manifest").Range("B2") = RouteName
    Worksheets("Manifest").Range("D5") = IBYardSlot

End Sub

'************************************************************************
'   PrintManifest
'
'   This sub will print the current contents of the Manifest worksheet
'
'   Inputs - None
'   Outputs - None
'************************************************************************
Sub PrintManifest()

    'Print Manifest
    Worksheets("Manifest").PrintOut

End Sub
