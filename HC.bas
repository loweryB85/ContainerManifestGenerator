Attribute VB_Name = "HC"
'Hide from User
Option Private Module

'Generate the helper column to assist with transferring identical BOLs onto manifest
Sub generateBOLHC()
    
    '**We must include the "With" statements in this module as opposed to in the calling module.**
    '**If placed in the calling module, then the Cells references in the code below would default to whatever the currently active worksheet is.**
    Dim Data As Worksheet
    
    Set Data = Worksheets("Data")
    
    With Data
    
        'BOL HC
        Dim current As Integer  'stores row of current BOL number being examined
        Dim previous As Integer 'stores row number of previous BOL number examined
        Dim HC As Integer
        
        'Define Column that contains BOL numbers to be used in creation of helper column
        Dim BOL_Column As Integer
        BOL_Column = 11
        
        'Define Column in which to create help column
        Dim HC_Column As Integer
        HC_Column = 13
    
        
        'Start at second line of Container data to skip header column - previous and current must begin as same value
        current = 2
        previous = current
        
        'iterate through list of BOL numbers until the currently examined cell is empty
        While (Not IsEmpty(.Cells(current, BOL_Column)))
        
            'Case 1
            'Special Case - first item in list
            If (current = previous) Then
                
                'Assign first BOL with initial HC value of 1
                .Cells(current, HC_Column).Value = 1
        
            'Case 2
            'Current BOL number is equal to the previous BOL number
            ElseIf (.Cells(current, BOL_Column).Value = .Cells(previous, BOL_Column).Value) Then
                
                'assign the same helper number given to previous BOL
                .Cells(current, HC_Column).Value = .Cells(previous, HC_Column).Value
            
            'Case 3
            'Current BOL number is not equal to the previous BOL number
            ElseIf (.Cells(current, BOL_Column).Value <> .Cells(previous, BOL_Column).Value) Then
            
                'assign helper number of 1 greater than previous helper number
                .Cells(current, HC_Column).Value = .Cells(previous, HC_Column).Value + 1
        
            Else
                    
                'Error message - this code should not execute if everything works properly
                .Cells(current, HC_Column).Value = "ERROR"
            
            End If
            
            'assign previous to current and increment current to continue moving through rows
            previous = current
            current = current + 1
            
        'end while loop
        Wend
        
    End With
    
End Sub

'Generate the helper column that will assist in transferring parts with identical destination docks to the manifest sheet
Sub generateDockHC()
    
    
    '**We must include the "With" statements in this module as opposed to in the calling module.**
    '**If placed in the calling module, then the Cells references in the code below would default to whatever the currently active worksheet is.**
    Dim Data As Worksheet
    
    Set Data = Worksheets("Data")
    
    With Data
    
        'Dock HC
        Dim current As Integer  'stores row of current BOL number being examined
        Dim previous As Integer 'stores row number of previous BOL number examined
        Dim HC As Integer
        
        'Define Column that contains BOL numbers to be used in creation of helper column
        Dim DOCK_Column As Integer
        DOCK_Column = 7
        
        'Define Column in which to create help column
        Dim HC_Column As Integer
        HC_Column = 14
    
        
        'Start at second line of Container data to skip header column - previous and current must begin as same value
        current = 2
        previous = current
        
        'iterate through list of BOL numbers until the currently examined cell is empty
        While (Not IsEmpty(.Cells(current, DOCK_Column)))
        
            'Case 1
            'Special Case - first item in list
            If (current = previous) Then
                
                'Assign first BOL with initial HC value of 1
                .Cells(current, HC_Column).Value = 1
        
            'Case 2
            'Current BOL number is equal to the previous BOL number
            ElseIf (.Cells(current, DOCK_Column).Value = .Cells(previous, DOCK_Column).Value) Then
                
                'assign the same helper number given to previous BOL
                .Cells(current, HC_Column).Value = .Cells(previous, HC_Column).Value
            
            'Case 3
            'Current BOL number is not equal to the previous BOL number
            ElseIf (.Cells(current, DOCK_Column).Value <> .Cells(previous, DOCK_Column).Value) Then
            
                'assign helper number of 1 greater than previous helper number
                .Cells(current, HC_Column).Value = .Cells(previous, HC_Column).Value + 1
        
            Else
                    
                'Error message - this code should not execute if everything works properly
                .Cells(current, HC_Column).Value = "ERROR"
            
            End If
            
            'assign previous to current and increment current to continue moving through rows
            previous = current
            current = current + 1
            
        'end while loop
        Wend
        
    End With
    
End Sub

