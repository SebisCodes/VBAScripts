'Insert this into your worksheet with the pivot table to trigger the function
'Private Sub Worksheet_SelectionChange(ByVal Target As range)
'  Application.EnableEvents = False
'  PivotTableDirectEdit.PivotTableDirectEdit 'Module.Sub
'  Application.EnableEvents = True
'End Sub

'Required References: mscorlib.dll, .NET Framework 3.5

' A VBA-script to direct edit pivot table values without changing to the source table
'
' Contact: sebiscodes@gmail.com
' My Github: https://github.com/SebisCodes

Sub PivotTableDirectEdit()
    On Error GoTo ErrorHandler
    'VALUE FILTER OPTION
    Dim activateValueFilterOption As Boolean: activateValueFilterOption = False
    
    'Vars for loops
    Dim c As Range
    Dim r As Range
    Dim i As Integer
    Dim v As Variant
    Dim tempField As PivotField
    Dim tempList As ArrayList
    
    'Get all pivot objects from the active cell
    Dim PT As PivotTable
    Dim PC As PivotCell
    Dim PF As PivotField
    Dim PI As PivotItem
    Set PT = ActiveCell.PivotTable
    Set PC = ActiveCell.PivotCell
    Set PF = ActiveCell.PivotField
    Set PI = ActiveCell.PivotItem
    
    Dim sourceDataName As String 'The rangename of the source data
    
    'Edit sourceData Name (from different languages to englisch, from R1C1 to A1)
    sourceDataName = PT.SourceData
    sourceDataName = ConvertLocalR1C1toR1C1(sourceDataName)
    sourceDataName = Application.ConvertFormula(sourceDataName, XlReferenceStyle.xlR1C1, XlReferenceStyle.xlA1)
    
    'Lists with column names and values to search for
    Dim allColNames As ArrayList: Set allColNames = New ArrayList 'All column names
    Dim dataColNames As ArrayList: Set dataColNames = New ArrayList 'Columns containing the dates
    Dim valueColNames As ArrayList: Set valueColNames = New ArrayList 'Columns containing the filter values
    Dim filterValues As ArrayList: Set filterValues = New ArrayList 'Values to search for
    
    'Store all columns (dates) in a list
    'Debug.Print "All Columns:"
    For Each tempField In PT.PivotFields
        If tempField.Name <> "Werte" Then
            allColNames.Add tempField.Name
            'Debug.Print " - " & tempField.Name
        End If
    Next
    
    'Store all data columns (dates) in a list
    'Debug.Print "All Data Columns:"
    For Each tempField In PT.DataFields
        dataColNames.Add tempField.SourceName
        'Debug.Print " - " & tempField.sourceName
    Next
    
    
    'Store all filter in a list
    'Debug.Print "All Filter Columns:"
    For Each v In allColNames
        If Not dataColNames.Contains(v) Then
            valueColNames.Add v
            'Debug.Print " - " & v
        Else
            Exit For
        End If
    Next
    
    'Store all filter values in a list
    'Debug.Print "All Values:"
    For Each v In PC.RowItems
        If v = "(blank)" Then
            v = ""
        End If
        filterValues.Add v
        'Debug.Print " - " & v
    Next
    For Each v In PC.ColumnItems
        If v = "(blank)" Then
            v = ""
        End If
        filterValues.Add v
        'Debug.Print " - " & v
    Next
    
    Dim colName As String: colName = PF.SourceName                      'The name of the selected column in the pivot table
    Dim sourceRange As Range: Set sourceRange = Range(sourceDataName)   'The range of all data in the source table
    Dim columnRange As Range                                            'The range of all columns in the source table
    'Get the first row
    For Each columnRange In sourceRange.Rows
        Exit For
    Next
    
    'Init some important vars
    Dim checkCounter As Integer
    Dim successfulRows As ArrayList: Set successfulRows = New ArrayList
    Dim foundRow As Integer
    Dim foundCol As Integer
    
    'Get Column
    For Each c In columnRange.Cells
        If c.Value2 = PF.SourceName Then
            foundCol = c.Column
            Exit For
        End If
    Next
    
    'Get Row (go row by row and check if all values of filterValues were found
    For Each r In sourceRange.Rows
        checkCounter = 0
        Set tempList = filterValues.Clone
        For Each c In r.Cells
            Call strInList(tempList, CStr(c.Value2), True)
            checkCounter = checkCounter + 1
            If checkCounter > valueColNames.Count - tempList.Count Then
                Exit For
            End If
            'Debug.Print successCounter
        Next
        If tempList.Count = 0 Then
            If activateValueFilterOption Then
                'VALUE FILTER OPTION
                If sourceRange.Worksheet.Cells(r.Row, foundCol).Value2 = ActiveCell.Value2 Then
                    successfulRows.Add r
                End If
            Else
                successfulRows.Add r
            End If
            If successfulRows.Count > 1 Then
                Exit For
            End If
        End If
    Next
    
    'Check if only one result was found
    If successfulRows.Count = 1 Then
        foundRow = successfulRows.Item(0).Row
    ElseIf successfulRows.Count = 0 Then
        Debug.Print "No data was found."
        Exit Sub
    Else
        Debug.Print "Please specify your cell more exactly. More than " & successfulRows.Count & " cells were found!"
        Exit Sub
    End If
    
    'Debug.Print foundRow & " - " & foundCol
    
    'Set val
    Dim newVal As String
    newVal = InputBox("Enter Value", "Change Value")
    
    If newVal <> "" Then
        If newVal = "0" Then newVal = ""
        sourceRange.Worksheet.Cells(foundRow, foundCol).Value2 = newVal
    End If
    
    'Refresh? If no - comment this
    PT.RefreshTable
    
ErrorHandler:
    
    
End Sub


'Check if str is in list
Function strInList(ByRef list As ArrayList, ByVal s As String, Optional ByVal removeElement As Boolean = False) As Boolean
    Dim i As Integer
    ret = False
    For i = 0 To list.Count - 1
        If list(i) = s Then
            ret = True
            If removeElement Then
                list.RemoveAt (i)
            End If
            Exit For
        End If
    Next
    strInList = ret
End Function

'Convert R1C1 from different languages to englisch to convert is after to A1
Function ConvertLocalR1C1toR1C1(localR1C1 As String) As String
    Dim result As String
    Dim language As Long
    language = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    
    Select Case language
        Case 1031 ' Deutsch
            result = Replace(localR1C1, "Z", "R")
            result = Replace(result, "S", "C")
        Case 1036 ' Französisch
            result = Replace(localR1C1, "L", "R")
            result = Replace(result, "C", "C")
        Case 3082 ' Spanisch
            result = Replace(localR1C1, "F", "R")
            result = Replace(result, "C", "C")
        Case Else
            result = localR1C1
    End Select
    
    ConvertLocalR1C1toR1C1 = result
End Function


