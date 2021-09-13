Attribute VB_Name = "ProcessManifest"
'Hide from User
Option Private Module

'************************************************************************
'   ProcessBills
'
'   This sub will process the parts in the STS container export by performing
'   a 3-level sort on the export data (BOL > Facility > Dock), then calling
'   another sub to generate a helper column to assist in iteration through the
'   bills, and finally iterating through each part in the export in order to
'   populate the manifest for each bill of lading.
'
'   Inputs - None
'   Outputs - Separate manifest for each bill of lading (printed)
'************************************************************************
Sub ProcessBills()

    Dim Data As Worksheet
    Dim toSort, ExportRangeEnd As Range
    Dim ExportNumRows, maxBills As Integer
    
    Set Data = Sheets("Data")
    
    With Data
        
        'get number of rows in STS export
        Set ExportRangeEnd = .Range("A" & .rows.count).End(xlUp)
        ExportNumRows = ExportRangeEnd.Row
        
        'For processing by ascending BOL numbers, sort by BOL, then Facility, then Dock
        'Order of sorting keys changed on 6/20/19 BL - this should ensure that identical BOL numbers, regardless of receiving facility/dock end up on the same manifests when printing by BOL
        Set toSort = .Range("A2:P" & ExportNumRows)
        toSort.Sort key1:=.Range("K2"), order1:=xlAscending, key2:=.Range("F2"), order2:=xlAscending, key3:=.Range("G2"), order3:=xlAscending
            
        'Now that the data has been sorted properly, we can generate the helper columns that will enable us to iterate through the list
        'in the proper order.
        HC.generateBOLHC
            
        maxBills = WorksheetFunction.Max(.Range("M:M"))
    
        'Begin looping through filtered data
        For count = 1 To maxBills
            
            ManifestFunctions.CopyPasteBOL (count)
            ManifestFunctions.PrintManifest
            Clear.ClearManifest
            
        Next count
        
    End With
    
    MsgBox ("Printing Complete")
    
End Sub

'************************************************************************
'   ProcessDocks
'
'   This sub will process the parts in the STS container export by performing
'   a 3-level sort on the export data (Facility > Dock > BOL), then calling
'   another sub to generate a helper column to assist in iteration through the
'   docks, and finally iterating through each part in the export in order to
'   populate the manifest for each receiving dock.
'
'   Inputs - None
'   Outputs - Separate manifest for each destination receiving dock (printed)
'************************************************************************
Sub ProcessDocks()

    Dim Data As Worksheet
    Dim toSort, ExportRangeEnd As Range
    Dim ExportNumRows, maxDocks As Integer
    
    Set Data = Sheets("Data")
    
    With Data
        
        'get number of rows in STS export
        Set ExportRangeEnd = .Range("A" & .rows.count).End(xlUp)
        ExportNumRows = ExportRangeEnd.Row
        
        'For processing by Dock, sort by Facility, then Dock, then BOL
        'Order of sorting keys changed on 6/20/19 BL - this should ensure that identical BOL numbers, regardless of receiving facility/dock end up on the same manifests when printing by BOL
        Set toSort = .Range("A2:P" & ExportNumRows)
        toSort.Sort key1:=.Range("F2"), order1:=xlAscending, key2:=.Range("G2"), order2:=xlAscending, key3:=.Range("K2"), order3:=xlAscending
            
        'Now that the data has been sorted properly, we can generate the helper columns that will enable us to iterate through the list
        'in the proper order.
        HC.generateDockHC
            
        maxDocks = WorksheetFunction.Max(.Range("N:N"))
    
        'Begin looping through filtered data
        For count = 1 To maxDocks
            
            ManifestFunctions.CopyPasteFacilityDock (count)
            ManifestFunctions.PrintManifest
            Clear.ClearManifest
            
        Next count
        
    End With

    MsgBox ("Printing Complete")

End Sub
