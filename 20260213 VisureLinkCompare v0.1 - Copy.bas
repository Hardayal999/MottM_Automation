Attribute VB_Name = "Module1"
Option Explicit

Sub VisureLinkCompare()
    'Main function to run all steps and compare tabs. This is ONLY a link compare, it does not compare changed text
    
    'SETUP VARIABLES BELOW
    '////////////////////////////////////////////////////////////////////////////////////////////////
    Dim currentBaselineTab As Worksheet
    Dim previousBaselineTab As Worksheet
    Set currentBaselineTab = ThisWorkbook.Sheets("Current")
    Set previousBaselineTab = ThisWorkbook.Sheets("Initialisation - PSTR 2.0 Impor")
    
    Dim mainCodeColumn As Integer
    mainCodeColumn = 1
    
    Dim lastRowCurrent As Long
    Dim lastRowPrevious As Long
    lastRowCurrent = currentBaselineTab.Cells(currentBaselineTab.Rows.Count, mainCodeColumn).End(xlUp).row
    lastRowPrevious = previousBaselineTab.Cells(previousBaselineTab.Rows.Count, mainCodeColumn).End(xlUp).row
    
    'define and assign the columns of linked Codes to compare
    Dim annexACodeColumn As Long
    annexACodeColumn = 9
    
    Dim annexBCodeColumn As Long
    annexBCodeColumn = 12
    
    Dim annexCCodeColumn As Long
    annexCCodeColumn = 15
    
    Dim annexDCodeColumn As Long
    annexDCodeColumn = 18
    
    Dim PCBCodeColumn As Long
    PCBCodeColumn = 21
    
    Dim endOfFile As Long
    endOfFile = 24
    
    Dim numberLinkedDocuments As Long
    numberLinkedDocuments = 5
    
    'This array should be 1 longer than "numberLinkedDocuments" above. Dont want to perform compare against the last number, this is there only to provide refernce for the
    Dim linkedDocColumns As Variant
    linkedDocColumns = Array(annexACodeColumn, annexBCodeColumn, annexCCodeColumn, annexDCodeColumn, PCBCodeColumn, endOfFile)
    
    '////////////////////////////////////////////////////////////////////////////////////////////////

        
    'Create new tab to store data
    Dim newCompareWorksheet As Worksheet
    Set newCompareWorksheet = CreateSheetWithUniqueName("COMPARE")
    
    'Compare tabs
    'Attain ID list in current and old tabs. This will need a comparison
    Dim firstRow As Long
    firstRow = 2
    Dim currentIDList, previousIDList, totalIDList As Variant
    
    currentIDList = ColumnToArray_UniqueMerged(currentBaselineTab.Range(currentBaselineTab.Cells(firstRow, mainCodeColumn), currentBaselineTab.Cells(lastRowCurrent, mainCodeColumn)), True, True)
    previousIDList = ColumnToArray_UniqueMerged(previousBaselineTab.Range(previousBaselineTab.Cells(firstRow, mainCodeColumn), previousBaselineTab.Cells(lastRowPrevious, mainCodeColumn)), True, True)
    
    'compare ID link arrays
    CompareArrays_Status2D currentIDList, previousIDList, totalIDList, False
    
    'attain all links to a given initial requirement. Count # links in current vs old. Create 3 arrays: one for each list of links, and one for the total results.
    'The final is a 2D array and stores a value that indicates whether text is added, the same, or removed. This will inform the final formatting
    'iteration starts here
    Dim newSheetRowCounter, codeCounter, columnCounter, totalMainIDCount As Long
    newSheetRowCounter = 2
    totalMainIDCount = UBound(totalIDList, 1) - LBound(totalIDList, 1) + 1 'INDEXING IS INCONSISTENT IN VBA. For totalIDList, first value is index 1
    
    'There is currently no test or error catching to see if totalMainIDCount = 0, we are just assuming it is not. It would go here however, same as the "If totalLinkIDCount > 1 Then" below if included
    
    'iterate through Codes first. INDEXING IS INCONSISTENT IN VBA. For totalIDList, first value is index 1
    For codeCounter = 1 To totalMainIDCount
        
        'flag to track scenario where all columns are empty of links (i.e. heading rows in export). If this is the case, there is a statement at the bottom of this for to paste the Main ID/Code just once.
        Dim noLinksFlag As Boolean
        noLinksFlag = True
        
        'Placed here so vlookup is not called every time in column and link loop. VLookupLike_RowSlice2D is a very poor solution overall with respect to runtime, and should be replaced in future. Will require a fundamental rebuild of the array approach.
        Dim mainCodePasteDetails As Variant
        If totalIDList(codeCounter, 2) = "Same" Or totalIDList(codeCounter, 2) = "Added" Then
            mainCodePasteDetails = VLookupLike_RowSlice2D(currentBaselineTab, _
                                                        mainCodeColumn, _
                                                        linkedDocColumns(0) - 1, _
                                                        2, _
                                                        lastRowCurrent, _
                                                        totalIDList(codeCounter, 1), _
                                                        caseSensitive:=True, _
                                                        trimMatchText:=True)
        ElseIf totalIDList(codeCounter, 2) = "Missing" Then
            mainCodePasteDetails = VLookupLike_RowSlice2D(previousBaselineTab, _
                                                        mainCodeColumn, _
                                                        linkedDocColumns(0) - 1, _
                                                        2, _
                                                        lastRowPrevious, _
                                                        totalIDList(codeCounter, 1), _
                                                        caseSensitive:=True, _
                                                        trimMatchText:=True)
        End If
        
        'iterate through Columns/Documents second. INDEXING IS INCONSISTENT IN VBA. For linkedDocColumns, first value is index 0
        For columnCounter = 0 To numberLinkedDocuments - 1 '-1 as first index is 0
                  
            Dim currentLinksToCompare, previousLinksToCompare, linkDifferences As Variant
            
            Call CollectSecondColumnValuesByIdentifier( _
                currentBaselineTab.Range(currentBaselineTab.Cells(firstRow, mainCodeColumn), currentBaselineTab.Cells(lastRowCurrent, mainCodeColumn)), _
                currentBaselineTab.Range(currentBaselineTab.Cells(firstRow, linkedDocColumns(columnCounter)), currentBaselineTab.Cells(lastRowCurrent, linkedDocColumns(columnCounter))), _
                totalIDList(codeCounter, 1), _
                currentLinksToCompare, _
                caseSensitive:=False, _
                trimMatchText:=True, _
                skipBlanks:=True)
            
            Call CollectSecondColumnValuesByIdentifier( _
                previousBaselineTab.Range(previousBaselineTab.Cells(firstRow, mainCodeColumn), previousBaselineTab.Cells(lastRowPrevious, mainCodeColumn)), _
                previousBaselineTab.Range(previousBaselineTab.Cells(firstRow, linkedDocColumns(columnCounter)), previousBaselineTab.Cells(lastRowPrevious, linkedDocColumns(columnCounter))), _
                totalIDList(codeCounter, 1), _
                previousLinksToCompare, _
                caseSensitive:=False, _
                trimMatchText:=True, _
                skipBlanks:=True)
            
            
            'compare link arrays from both tabs
            CompareArrays_Status2D currentLinksToCompare, previousLinksToCompare, linkDifferences, False
            
            'iterate through links third
            'paste array into the created tab with formatting
            Dim linkCounter, totalLinkIDCount As Long
            totalLinkIDCount = UBound(linkDifferences, 1) - LBound(linkDifferences, 1) + 1
            
            'INDEXING IS INCONSISTENT IN VBA. For linkDifferences, first value is index 1
            If totalLinkIDCount > 1 Then
                
                'Have at least one link. No need for final paste command
                noLinksFlag = False
                                          
                For linkCounter = 1 To totalLinkIDCount
                    
                    'Check if "Added" | "Same" | "Missing". Checks the main Code match first, and then the link Code match second
                    If totalIDList(codeCounter, 2) = "Same" Then
                        'expected result, perform second check on link Code match
                        
                        If linkDifferences(linkCounter, 2) = "Same" Then
                            'black - no formatting required. Link is retained
                            'paste command. Pastes Main ID in every row, and link id
                            newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Value = totalIDList(codeCounter, 1)
                            newCompareWorksheet.Cells(newSheetRowCounter, linkedDocColumns(columnCounter)).Value = linkDifferences(linkCounter, 1)
                            
                            'paste main ID details
                            newCompareWorksheet.Range(Cells(newSheetRowCounter, mainCodeColumn + 1), Cells(newSheetRowCounter, linkedDocColumns(0) - 1)).Value = mainCodePasteDetails
                            'paste link ID details. If same, search Current tab. VLookupLike_RowSlice2D is a very poor solution overall with respect to runtime, and should be replaced in future. Will require a fundamental rebuild of the array approach.
                            newCompareWorksheet.Range(Cells(newSheetRowCounter, linkedDocColumns(columnCounter) + 1), Cells(newSheetRowCounter, _
                                                                linkedDocColumns(columnCounter + 1) - 1)).Value = VLookupLike_RowSlice2D(currentBaselineTab, _
                                                                                                                                linkedDocColumns(columnCounter), _
                                                                                                                                linkedDocColumns(columnCounter + 1) - 1, _
                                                                                                                                2, _
                                                                                                                                lastRowCurrent, _
                                                                                                                                linkDifferences(linkCounter, 1), _
                                                                                                                                caseSensitive:=True, _
                                                                                                                                trimMatchText:=True)
                            
                        ElseIf linkDifferences(linkCounter, 2) = "Added" Then
                            'green formatting. Link ID is new
                            'paste command. Pastes Main ID in every row, and link id
                            newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Value = totalIDList(codeCounter, 1)
                            newCompareWorksheet.Cells(newSheetRowCounter, linkedDocColumns(columnCounter)).Value = linkDifferences(linkCounter, 1)
                            
                            'paste main ID details
                            newCompareWorksheet.Range(Cells(newSheetRowCounter, mainCodeColumn + 1), Cells(newSheetRowCounter, linkedDocColumns(0) - 1)).Value = mainCodePasteDetails
                            'paste link ID details. If added, search current tab. VLookupLike_RowSlice2D is a very poor solution overall with respect to runtime, and should be replaced in future. Will require a fundamental rebuild of the array approach.
                            newCompareWorksheet.Range(Cells(newSheetRowCounter, linkedDocColumns(columnCounter) + 1), Cells(newSheetRowCounter, _
                                                                linkedDocColumns(columnCounter + 1) - 1)).Value = VLookupLike_RowSlice2D(currentBaselineTab, _
                                                                                                                                linkedDocColumns(columnCounter), _
                                                                                                                                linkedDocColumns(columnCounter + 1) - 1, _
                                                                                                                                2, _
                                                                                                                                lastRowCurrent, _
                                                                                                                                linkDifferences(linkCounter, 1), _
                                                                                                                                caseSensitive:=True, _
                                                                                                                                trimMatchText:=True)
                            
                            'format command. Just format Link ID
                            newCompareWorksheet.Cells(newSheetRowCounter, linkedDocColumns(columnCounter)).Font.Color = vbGreen
                            newCompareWorksheet.Range(Cells(newSheetRowCounter, linkedDocColumns(columnCounter) + 1), Cells(newSheetRowCounter, _
                                                                linkedDocColumns(columnCounter + 1) - 1)).Font.Color = vbGreen
                            
                        ElseIf linkDifferences(linkCounter, 2) = "Missing" Then
                            'red formatting. Link ID removed
                            'paste command. Pastes Main ID in every row, and link id
                            newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Value = totalIDList(codeCounter, 1)
                            newCompareWorksheet.Cells(newSheetRowCounter, linkedDocColumns(columnCounter)).Value = linkDifferences(linkCounter, 1)
                            
                            'paste main ID details
                            newCompareWorksheet.Range(Cells(newSheetRowCounter, mainCodeColumn + 1), Cells(newSheetRowCounter, linkedDocColumns(0) - 1)).Value = mainCodePasteDetails
                            'paste link ID details. If added, search Previous tab. VLookupLike_RowSlice2D is a very poor solution overall with respect to runtime, and should be replaced in future. Will require a fundamental rebuild of the array approach.
                            newCompareWorksheet.Range(Cells(newSheetRowCounter, linkedDocColumns(columnCounter) + 1), Cells(newSheetRowCounter, _
                                                                linkedDocColumns(columnCounter + 1) - 1)).Value = VLookupLike_RowSlice2D(previousBaselineTab, _
                                                                                                                                linkedDocColumns(columnCounter), _
                                                                                                                                linkedDocColumns(columnCounter + 1) - 1, _
                                                                                                                                2, _
                                                                                                                                lastRowPrevious, _
                                                                                                                                linkDifferences(linkCounter, 1), _
                                                                                                                                caseSensitive:=True, _
                                                                                                                                trimMatchText:=True)
                            
                            'format command. Just format Link ID
                            newCompareWorksheet.Cells(newSheetRowCounter, linkedDocColumns(columnCounter)).Font.Color = vbRed
                            newCompareWorksheet.Range(Cells(newSheetRowCounter, linkedDocColumns(columnCounter) + 1), Cells(newSheetRowCounter, _
                                                                linkedDocColumns(columnCounter + 1) - 1)).Font.Color = vbRed
                            
                        Else
                            'Fail
                            'Add line to write into the new sheet that an ID was skipped. Print the ID
                            'paste command. Pastes Main ID in every row, and link id
                            newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Value = totalIDList(codeCounter, 1)
                            newCompareWorksheet.Cells(newSheetRowCounter, linkedDocColumns(columnCounter)).Value = "FAILED: <" & linkDifferences(linkCounter, 1) & ">"
                            
                            'format command. Just format Link ID
                            newCompareWorksheet.Cells(newSheetRowCounter, linkedDocColumns(columnCounter)).Interior.Color = vbYellow
                            
                        End If
                        
                    ElseIf totalIDList(codeCounter, 2) = "Added" Then
                        'green formatting for entire row. Main ID is new
                        'paste command. Pastes Main ID in every row, and link id
                        newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Value = totalIDList(codeCounter, 1)
                        newCompareWorksheet.Cells(newSheetRowCounter, linkedDocColumns(columnCounter)).Value = linkDifferences(linkCounter, 1)
                        
                        'paste main ID details
                        newCompareWorksheet.Range(Cells(newSheetRowCounter, mainCodeColumn + 1), Cells(newSheetRowCounter, linkedDocColumns(0) - 1)).Value = mainCodePasteDetails
                        'paste link ID details. If added, search current tab. VLookupLike_RowSlice2D is a very poor solution overall with respect to runtime, and should be replaced in future. Will require a fundamental rebuild of the array approach.
                        newCompareWorksheet.Range(Cells(newSheetRowCounter, linkedDocColumns(columnCounter) + 1), Cells(newSheetRowCounter, _
                                                            linkedDocColumns(columnCounter + 1) - 1)).Value = VLookupLike_RowSlice2D(currentBaselineTab, _
                                                                                                                            linkedDocColumns(columnCounter), _
                                                                                                                            linkedDocColumns(columnCounter + 1) - 1, _
                                                                                                                            2, _
                                                                                                                            lastRowCurrent, _
                                                                                                                            linkDifferences(linkCounter, 1), _
                                                                                                                            caseSensitive:=True, _
                                                                                                                            trimMatchText:=True)
                        
                        'format command
                        newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Font.Color = vbGreen
                        newCompareWorksheet.Cells(newSheetRowCounter, linkedDocColumns(columnCounter)).Font.Color = vbGreen
                        
                        newCompareWorksheet.Range(Cells(newSheetRowCounter, mainCodeColumn + 1), Cells(newSheetRowCounter, linkedDocColumns(0) - 1)).Font.Color = vbGreen
                        newCompareWorksheet.Range(Cells(newSheetRowCounter, linkedDocColumns(columnCounter) + 1), Cells(newSheetRowCounter, _
                                                            linkedDocColumns(columnCounter + 1) - 1)).Font.Color = vbGreen
                        
                    ElseIf totalIDList(codeCounter, 2) = "Missing" Then
                        'red formatting for entire row. Main ID removed
                        'paste command. Pastes Main ID in every row, and link id
                        newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Value = totalIDList(codeCounter, 1)
                        newCompareWorksheet.Cells(newSheetRowCounter, linkedDocColumns(columnCounter)).Value = linkDifferences(linkCounter, 1)
                        
                        'paste main ID details
                        newCompareWorksheet.Range(Cells(newSheetRowCounter, mainCodeColumn + 1), Cells(newSheetRowCounter, linkedDocColumns(0) - 1)).Value = mainCodePasteDetails
                        'paste link ID details. If added, search Previous tab. VLookupLike_RowSlice2D is a very poor solution overall with respect to runtime, and should be replaced in future. Will require a fundamental rebuild of the array approach.
                        newCompareWorksheet.Range(Cells(newSheetRowCounter, linkedDocColumns(columnCounter) + 1), Cells(newSheetRowCounter, _
                                                            linkedDocColumns(columnCounter + 1) - 1)).Value = VLookupLike_RowSlice2D(previousBaselineTab, _
                                                                                                                            linkedDocColumns(columnCounter), _
                                                                                                                            linkedDocColumns(columnCounter + 1) - 1, _
                                                                                                                            2, _
                                                                                                                            lastRowPrevious, _
                                                                                                                            linkDifferences(linkCounter, 1), _
                                                                                                                            caseSensitive:=True, _
                                                                                                                            trimMatchText:=True)
                        
                        'format command. Just format both Codes/IDs
                        newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Font.Color = vbRed
                        newCompareWorksheet.Cells(newSheetRowCounter, linkedDocColumns(columnCounter)).Font.Color = vbRed
                        
                        newCompareWorksheet.Range(Cells(newSheetRowCounter, mainCodeColumn + 1), Cells(newSheetRowCounter, linkedDocColumns(0) - 1)).Font.Color = vbRed
                        newCompareWorksheet.Range(Cells(newSheetRowCounter, linkedDocColumns(columnCounter) + 1), Cells(newSheetRowCounter, _
                                                            linkedDocColumns(columnCounter + 1) - 1)).Font.Color = vbRed
                        
                    Else
                        'Fail
                        'Add line to write into the new sheet that an ID was skipped. Print the ID
                        'paste command. Pastes Main ID in every row, and link id
                        newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Value = totalIDList(codeCounter, 1)
                        newCompareWorksheet.Cells(newSheetRowCounter, linkedDocColumns(columnCounter)).Value = "FAILED: <" & linkDifferences(linkCounter, 1) & ">"
                        
                        'format command. Just format Link ID
                        newCompareWorksheet.Cells(newSheetRowCounter, linkedDocColumns(columnCounter)).Interior.Color = vbYellow
                    End If
                
                    'MOVE SHEET COUNTER
                    newSheetRowCounter = newSheetRowCounter + 1
                    
                    'repeat for next link ID/Code
                    
                Next
                
            Else
                'linkDifferences is empty, do nothing
            End If
            
            'repeat for next column
        
        Next
        
        'If all columns were empty of links (i.e. headings), just paste the ID once.
        If noLinksFlag Then
            'Check if "Added" | "Same" | "Missing". Checks the main Code match first, and then the link Code match second
            If totalIDList(codeCounter, 2) = "Same" Then
                'black, no formatting
                newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Value = totalIDList(codeCounter, 1)
                'paste main ID details
                newCompareWorksheet.Range(Cells(newSheetRowCounter, mainCodeColumn + 1), Cells(newSheetRowCounter, linkedDocColumns(0) - 1)).Value = mainCodePasteDetails
                            
                
            ElseIf totalIDList(codeCounter, 2) = "Added" Then
                'green formatting for entire row. Main ID is new
                newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Value = totalIDList(codeCounter, 1)
                'paste main ID details
                newCompareWorksheet.Range(Cells(newSheetRowCounter, mainCodeColumn + 1), Cells(newSheetRowCounter, linkedDocColumns(0) - 1)).Value = mainCodePasteDetails
                
                'format command.
                newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Font.Color = vbGreen
                newCompareWorksheet.Range(Cells(newSheetRowCounter, mainCodeColumn + 1), Cells(newSheetRowCounter, linkedDocColumns(0) - 1)).Font.Color = vbGreen
                
            ElseIf totalIDList(codeCounter, 2) = "Missing" Then
                'red formatting for entire row. Main ID removed
                newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Value = totalIDList(codeCounter, 1)
                'paste main ID details
                newCompareWorksheet.Range(Cells(newSheetRowCounter, mainCodeColumn + 1), Cells(newSheetRowCounter, linkedDocColumns(0) - 1)).Value = mainCodePasteDetails
                                
                'format command.
                newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Font.Color = vbRed
                newCompareWorksheet.Range(Cells(newSheetRowCounter, mainCodeColumn + 1), Cells(newSheetRowCounter, linkedDocColumns(0) - 1)).Font.Color = vbRed
                
            Else
                'Fail
                'Add line to write into the new sheet that an ID was skipped. Print the ID
                'paste command. Pastes Main ID in every row, and link id
                newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Value = "FAILED: <" & totalIDList(codeCounter, 1) & ">"
                
                'format command
                newCompareWorksheet.Cells(newSheetRowCounter, mainCodeColumn).Interior.Color = vbYellow
                
            End If
            
            'MOVE SHEET COUNTER
            newSheetRowCounter = newSheetRowCounter + 1
        End If
        
        'repeat for next ID
        
    Next
    
    
    'dummy line to use as stop point
    Dim stoppoint As Long
    stoppoint = 1
    
End Sub

Sub Example_CompareArrays()
    Dim previousArray As Variant
    Dim currentArray As Variant
    Dim result As Variant
    Dim i As Long
    Dim j As Long
    
    previousArray = Array("Alpha", "Bravo", "Charlie", "Echo")
    currentArray = Array("Bravo", "Delta", "Echo", "Foxtrot")
    
    ' caseSensitive:=False ? case-insensitive for strings
    CompareArrays_Status2D currentArray, previousArray, result, False
    
    ' Print results
    For i = LBound(result, 1) To UBound(result, 1)
        For j = 1 To 2
            Debug.Print result(i, j)
        Next
    Next
End Sub

' Compares two arrays (current vs previous) as sets and returns a 2D [N x 2] Variant array:
'   Col 1: Item
'   Col 2: "Added" | "Same" | "Missing"
'
' - Treats arrays as sets (duplicates collapsed).
' - Order: items from previous first (in their original order), then newly added items from current.
' - Works with strings, numbers, dates, booleans. Case sensitivity is configurable.
'
' Example call:
'   Dim out As Variant
'   CompareArrays_Status2D currentArray, previousArray, out, False
'   ' out is 1..N x 1..2
'
Public Sub CompareArrays_Status2D( _
    ByVal currentArray As Variant, _
    ByVal previousArray As Variant, _
    ByRef result As Variant, _
    Optional ByVal caseSensitive As Boolean = False)

    Dim dPrev As Object, dCurr As Object, seen As Object
    Dim order As Collection
    Dim arrPrev As Variant, arrCurr As Variant
    Dim i As Long, key As Variant, n As Long
    Dim cmpMode As Long
    
    ' Choose comparison mode for strings
    cmpMode = IIf(caseSensitive, vbBinaryCompare, vbTextCompare)
    
    ' Normalize inputs to 1D arrays
    arrPrev = ToFlat1D(previousArray)
    arrCurr = ToFlat1D(currentArray)
    
    ' Dictionaries to record membership
    Set dPrev = CreateObject("Scripting.Dictionary")
    Set dCurr = CreateObject("Scripting.Dictionary")
    dPrev.CompareMode = cmpMode
    dCurr.CompareMode = cmpMode
    
    ' Fill previous membership (preserve discovery order via Collection)
    Set order = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = cmpMode
    
    If Not IsEmpty1D(arrPrev) Then
        For i = LBound(arrPrev) To UBound(arrPrev)
            key = arrPrev(i)
            If Not dPrev.Exists(key) Then dPrev(key) = True
            If Not seen.Exists(key) Then
                order.Add key
                seen(key) = True
            End If
        Next i
    End If
    
    ' Fill current membership; append to order any items not yet seen
    If Not IsEmpty1D(arrCurr) Then
        For i = LBound(arrCurr) To UBound(arrCurr)
            key = arrCurr(i)
            If Not dCurr.Exists(key) Then dCurr(key) = True
            If Not seen.Exists(key) Then
                order.Add key
                seen(key) = True
            End If
        Next i
    End If
    
    ' Build output [N x 2]
    n = order.Count
    If n = 0 Then
        ' Return a zero-length 2D array
        'Old code - buggy. Can't get this to work - just return an empty 1D array. Need to add a line in main function to check if this is returning an empty array.
        'ReDim result(1 To 0, 1 To 2)
        
        result = VBA.Array(VBA.Array()) ' zero-length
        
        Exit Sub
    End If
    
    ReDim result(1 To n, 1 To 2)
    For i = 1 To n
        key = order(i)
        result(i, 1) = key
        If dPrev.Exists(key) And dCurr.Exists(key) Then
            result(i, 2) = "Same"
        ElseIf dCurr.Exists(key) Then
            result(i, 2) = "Added"
        Else
            result(i, 2) = "Missing"
        End If
    Next i
End Sub

' ----------------------------
' Helpers
' ----------------------------

' Flattens a variety of inputs into a 1D Variant array.
' Accepts:
'   - 1D VBA arrays
'   - 2D arrays (e.g., from Range.Value); flattens row-major
'   - A single value (wraps as 1-element array)
'   - Range (row or column or block); flattens row-major
Private Function ToFlat1D(ByVal v As Variant) As Variant
    Dim arr As Variant
    Dim r As Long, c As Long, i As Long, j As Long, n As Long
    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
    Dim tmp() As Variant
    
    If IsObject(v) Then
        If TypeOf v Is Excel.Range Then
            arr = v.Value
            If IsArray(arr) Then
                ' 2D range -> flatten
                lb1 = LBound(arr, 1): ub1 = UBound(arr, 1)
                lb2 = LBound(arr, 2): ub2 = UBound(arr, 2)
                ReDim tmp(0 To (ub1 - lb1 + 1) * (ub2 - lb2 + 1) - 1)
                n = 0
                For i = lb1 To ub1
                    For j = lb2 To ub2
                        tmp(n) = arr(i, j)
                        n = n + 1
                    Next j
                Next i
                ToFlat1D = tmp
                Exit Function
            Else
                ' Single cell
                ToFlat1D = VBA.Array(arr)
                Exit Function
            End If
        End If
    End If
    
    If IsArray(v) Then
        On Error GoTo oneDim
        ' Try 2D
        lb1 = LBound(v, 1): ub1 = UBound(v, 1)
        lb2 = LBound(v, 2): ub2 = UBound(v, 2)
        ReDim tmp(0 To (ub1 - lb1 + 1) * (ub2 - lb2 + 1) - 1)
        n = 0
        For i = lb1 To ub1
            For j = lb2 To ub2
                tmp(n) = v(i, j)
                n = n + 1
            Next j
        Next i
        ToFlat1D = tmp
        Exit Function
oneDim:
        ' 1D array
        Err.Clear
        ToFlat1D = v
        Exit Function
    End If
    
    ' Non-array scalar
    ToFlat1D = VBA.Array(v)
End Function

' Detects a zero-length or uninitialized 1D array
Private Function IsEmpty1D(ByVal a As Variant) As Boolean
    On Error GoTo emptyState
    If Not IsArray(a) Then IsEmpty1D = True: Exit Function
    If UBound(a) < LBound(a) Then IsEmpty1D = True Else IsEmpty1D = False
    Exit Function
emptyState:
    IsEmpty1D = True
End Function


' Creates a new Worksheet with a unique name derived from baseName.
' If "COMPARE" exists, tries "COMPARE1", "COMPARE2", ...
' Returns the newly created Worksheet and activates it.
'
' Parameters:
'   baseName         - the desired base sheet name (e.g., "COMPARE")
'   insertAtEnd      - if True, insert after the last sheet; if False, insert before the first
'   tryLimit         - optional safety cap on attempts (default 1000)
'
Public Function CreateSheetWithUniqueName( _
    ByVal baseName As String, _
    Optional ByVal insertAtEnd As Boolean = True, _
    Optional ByVal tryLimit As Long = 1000) As Worksheet

    Dim nameToTry As String
    Dim suffix As Long
    Dim newWS As Worksheet
    Dim safeBase As String
    
    If Len(baseName) = 0 Then Err.Raise vbObjectError + 1, , "Base name cannot be empty."
    
    ' Ensure the base name itself is valid and within 31 characters.
    safeBase = MakeValidSheetName(baseName)
    If Len(safeBase) = 0 Then Err.Raise vbObjectError + 2, , "Base name is not valid for a sheet."
    
    ' First try the base name as-is
    nameToTry = TruncateForSuffix(safeBase, 0)
    If Not SheetNameExists(nameToTry) Then
        Set newWS = AddSheet(nameToTry, insertAtEnd)
        Set CreateSheetWithUniqueName = newWS
        Exit Function
    End If
    
    ' Otherwise append numeric suffixes: 1, 2, 3, ...
    For suffix = 1 To tryLimit
        nameToTry = TruncateForSuffix(safeBase, suffix)
        If Not SheetNameExists(nameToTry) Then
            Set newWS = AddSheet(nameToTry, insertAtEnd)
            Set CreateSheetWithUniqueName = newWS
            Exit Function
        End If
    Next suffix
    
    Err.Raise vbObjectError + 3, , "Unable to find a unique sheet name after " & tryLimit & " attempts."
End Function

' Adds a worksheet with the given name, either at the end or the beginning.
Private Function AddSheet(ByVal sheetName As String, ByVal insertAtEnd As Boolean) As Worksheet
    Dim ws As Worksheet
    If insertAtEnd Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    Else
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    End If
    ws.Name = sheetName
    ws.Activate
    Set AddSheet = ws
End Function

' Returns True if any sheet (Worksheet/Charts/etc.) in the workbook already has this name.
Private Function SheetNameExists(ByVal sheetName As String) As Boolean
    Dim s As Object
    For Each s In ThisWorkbook.Sheets
        If StrComp(s.Name, sheetName, vbBinaryCompare) = 0 Then
            SheetNameExists = True
            Exit Function
        End If
    Next s
    SheetNameExists = False
End Function

' Ensures name is valid: removes invalid characters and trims whitespace.
' Invalid characters: : \ / ? * [ ] and leading/trailing apostrophes or spaces.
' Also collapses repeated spaces.
Private Function MakeValidSheetName(ByVal nm As String) As String
    Dim badChars As Variant, ch As Variant
    nm = Trim$(nm)
    
    ' Remove invalid characters
    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    For Each ch In badChars
        nm = Replace$(nm, CStr(ch), "")
    Next ch
    
    ' Excel prohibits names that start or end with apostrophe
    Do While Left$(nm, 1) = "'"
        nm = Mid$(nm, 2)
    Loop
    Do While Right$(nm, 1) = "'"
        nm = Left$(nm, Len(nm) - 1)
    Loop
    
    ' Collapse repeated spaces
    Do While InStr(nm, "  ") > 0
        nm = Replace$(nm, "  ", " ")
    Loop
    
    ' Truncate to 31 if too long
    If Len(nm) > 31 Then nm = Left$(nm, 31)
    
    MakeValidSheetName = nm
End Function

' Truncates/adjusts base so that base & CStr(suffix) fits within 31 characters.
' If suffix = 0, returns base (already truncated by MakeValidSheetName).
Private Function TruncateForSuffix(ByVal baseName As String, ByVal suffix As Long) As String
    Const MAXLEN As Long = 31
    Dim sfx As String
    Dim needed As Long
    Dim trimmedBase As String
    
    If suffix = 0 Then
        TruncateForSuffix = Left$(baseName, MAXLEN)
        Exit Function
    End If
    
    sfx = CStr(suffix)
    needed = Len(baseName) + Len(sfx)
    
    If needed <= MAXLEN Then
        TruncateForSuffix = baseName & sfx
    Else
        trimmedBase = Left$(baseName, MAXLEN - Len(sfx))
        TruncateForSuffix = trimmedBase & sfx
    End If
End Function


' Converts a single-column range to a 1D Variant array, taking
' only the first row of each merged area (skips the subsequent rows in the same merge).
' Optional: Trim strings and skip blanks.
Public Function ColumnToArray_UniqueMerged( _
    ByVal colRange As Range, _
    Optional ByVal trimStrings As Boolean = True, _
    Optional ByVal skipBlanks As Boolean = False _
) As Variant
    
    Dim cell As Range, ma As Range
    Dim out() As Variant
    Dim i As Long, v As Variant, s As String
    
    If colRange Is Nothing Then
        ReDim out(0 To -1): ColumnToArray_UniqueMerged = out: Exit Function
    End If
    If colRange.Columns.Count <> 1 Then Err.Raise vbObjectError + 1, , "Range must be a single column."
    
    ' Worst case: every row yields one output
    ReDim out(0 To colRange.Rows.Count - 1)
    i = 0
    
    For Each cell In colRange.Cells
        If cell.MergeCells Then
            Set ma = cell.MergeArea
            ' Take only when we're at the first cell of the merge area
            If cell.Address = ma.Cells(1, 1).Address Then
                v = ma.Cells(1, 1).Value
                If trimStrings And VarType(v) = vbString Then v = Trim$(CStr(v))
                
                If skipBlanks Then
                    s = CStr(v)
                    If LenB(s) > 0 Then out(i) = v: i = i + 1
                Else
                    out(i) = v: i = i + 1
                End If
            End If
        Else
            v = cell.Value
            If trimStrings And VarType(v) = vbString Then v = Trim$(CStr(v))
            
            If skipBlanks Then
                s = CStr(v)
                If LenB(s) > 0 Then out(i) = v: i = i + 1
            Else
                out(i) = v: i = i + 1
            End If
        End If
    Next cell
    
    ' Shrink to actual size
    If i = 0 Then
        ReDim out(0 To -1)
    ElseIf i <= UBound(out) Then
        ReDim Preserve out(0 To i - 1)
    End If
    
    ColumnToArray_UniqueMerged = out
End Function


' Collects values from secondCol for rows where firstCol equals `identifier`.
' - Handles merged cells in firstCol: if the merged area's top-left cell matches the identifier,
'   ALL rows covered by that merged area are included.
' - Skips blanks/empty/whitespace values from secondCol.
' - Outputs a 1-D Variant array via ByRef `result`.
'
' Parameters:
'   firstCol        Single-column Range to match against (e.g., Sheet1.Range("A2:A500"))
'   secondCol       Single-column Range aligned row-for-row with firstCol (e.g., Sheet1.Range("B2:B500"))
'   identifier      Text to match against values in firstCol
'   result          [ByRef] Receives a 1-D Variant array of collected values (0-length if none)
'   caseSensitive   Optional; default False (case-insensitive string match)
'   trimMatchText   Optional; default True (trims both sides before comparing)
'   skipBlanks      Optional; default True (removes blanks/empty/whitespace from result)
'
Public Sub CollectSecondColumnValuesByIdentifier( _
    ByVal firstCol As Range, _
    ByVal secondCol As Range, _
    ByVal identifier As String, _
    ByRef result As Variant, _
    Optional ByVal caseSensitive As Boolean = False, _
    Optional ByVal trimMatchText As Boolean = True, _
    Optional ByVal skipBlanks As Boolean = True _
)
    Dim rowCount As Long
    Dim r As Long, idx As Long
    Dim top As Long
    Dim firstCell As Range, ma As Range, block As Range
    Dim arrSecond As Variant
    Dim out() As Variant
    Dim idNorm As String
    Dim v As Variant
    Dim blockFirstIdx As Long, blockLastIdx As Long
    Dim cmpOk As Boolean
    
    ' ---- Validation ----
    If firstCol Is Nothing Or secondCol Is Nothing Then
        Err.Raise vbObjectError + 101, , "Both firstCol and secondCol must be provided."
    End If
    If firstCol.Columns.Count <> 1 Or secondCol.Columns.Count <> 1 Then
        Err.Raise vbObjectError + 102, , "Both firstCol and secondCol must be single columns."
    End If
    If firstCol.Rows.Count <> secondCol.Rows.Count Then
        Err.Raise vbObjectError + 103, , "firstCol and secondCol must have the same number of rows."
    End If
    If Not firstCol.Worksheet Is secondCol.Worksheet Then
        Err.Raise vbObjectError + 104, , "firstCol and secondCol must be on the same worksheet."
    End If
    
    rowCount = firstCol.Rows.Count
    top = firstCol.row
    
    ' Normalize identifier once
    idNorm = identifier
    If trimMatchText Then idNorm = Trim$(idNorm)
    
    ' Load second column values at once (fast)
    arrSecond = secondCol.Value  ' 2D: (1..rowCount, 1..1)
    
    ' Pre-size output to worst case (every row matches), then trim
    ReDim out(1 To rowCount)
    idx = 0
    
    ' Iterate rows with a Do loop so we can jump over merged blocks
    r = 1
    Do While r <= rowCount
        Set firstCell = firstCol.Cells(r, 1)
        
        If firstCell.MergeCells Then
            Set ma = firstCell.MergeArea
            ' Limit to the part of merge area that intersects our firstCol range
            Set block = Intersect(ma, firstCol)
            If Not block Is Nothing Then
                ' Compute indices within our ranges
                blockFirstIdx = block.row - top + 1
                blockLastIdx = blockFirstIdx + block.Rows.Count - 1
                
                ' Compare once using the top-left value of the merge area
                v = ma.Cells(1, 1).Value
                cmpOk = ValuesEqualToIdentifier(v, idNorm, caseSensitive, trimMatchText)
                
                If cmpOk Then
                    ' Collect from second col for all rows in the block
                    Dim k As Long, val2 As Variant
                    For k = blockFirstIdx To blockLastIdx
                        val2 = arrSecond(k, 1)
                        If Not (skipBlanks And IsBlankLike(val2)) Then
                            idx = idx + 1
                            out(idx) = val2
                        End If
                    Next k
                End If
                
                ' Skip to the row after the merged block
                r = blockLastIdx + 1
            Else
                ' Shouldn't happen, but advance safely
                r = r + 1
            End If
        Else
            ' Non-merged: compare value in this row
            v = firstCell.Value
            If ValuesEqualToIdentifier(v, idNorm, caseSensitive, trimMatchText) Then
                Dim val1 As Variant
                val1 = arrSecond(r, 1)
                If Not (skipBlanks And IsBlankLike(val1)) Then
                    idx = idx + 1
                    out(idx) = val1
                End If
            End If
            r = r + 1
        End If
    Loop
    
    ' Trim output to actual size or return a 0-length array
    If idx = 0 Then
        result = VBA.Array() 'ReDim result(1 To 0)  ' zero-length
    Else
        ReDim Preserve out(1 To idx)
        result = out
    End If
End Sub

' --- Helpers ---

' Compares a cell's value (Variant) to the identifier string using string semantics.
' Converts the cell value to string; optionally trims both sides; then uses StrComp with chosen sensitivity.
Private Function ValuesEqualToIdentifier( _
    ByVal cellValue As Variant, _
    ByVal idNorm As String, _
    ByVal caseSensitive As Boolean, _
    ByVal trimMatchText As Boolean _
) As Boolean
    Dim s As String
    s = CStr(cellValue)
    If trimMatchText Then s = Trim$(s)
    
    If caseSensitive Then
        ValuesEqualToIdentifier = (StrComp(s, idNorm, vbBinaryCompare) = 0)
    Else
        ValuesEqualToIdentifier = (StrComp(s, idNorm, vbTextCompare) = 0)
    End If
End Function

' Treats Variant as blank if:
'   - IsEmpty, or
'   - IsNull, or
'   - Is a zero-length/whitespace-only String.
' Numbers/dates/booleans are considered non-blank.
Private Function IsBlankLike(ByVal v As Variant) As Boolean
    If IsError(v) Then
        ' Treat error values as non-blank; change to True if you prefer to skip errors
        IsBlankLike = False
        Exit Function
    End If
    If IsEmpty(v) Then IsBlankLike = True: Exit Function
    If IsNull(v) Then IsBlankLike = True: Exit Function
    If VarType(v) = vbString Then
        IsBlankLike = (LenB(Trim$(CStr(v))) = 0)
    Else
        IsBlankLike = False
    End If
End Function


' Returns a 1 x (endCol - startCol) 2-D array of strings for the first row
' where ws.Cells(row, startCol) matches `identifier` (first match only).
' The returned slice EXCLUDES the identifier column, i.e., it is the row values
' from columns startCol+1 .. endCol (all coerced to String).
'
' Parameters:
'   ws               - Worksheet containing the data.
'   startCol         - First column number (e.g., 1 for column A). This is the column to search.
'   endCol           - Last column number to include in the return (inclusive).
'   firstDataRow     - First row to search (e.g., 2 if headers are on row 1).
'   lastDataRow      - Last row to search.
'   identifier       - The lookup text to match (treated as string).
'   caseSensitive    - Optional; default False (case-insensitive).
'   trimMatchText    - Optional; default True (Trim both cell & identifier before comparing).
'   handleMergedKeys - Optional; default True (treat a merged block's top-left as the key).
'   matchedRow       - Optional [ByRef]; returns the row index where the match was found (0 if not found).
'
' Returns:
'   Variant 2-D array sized (1 to 1, 1 to endCol - startCol). If no match, returns the same
'   sized 2-D array prefilled with empty strings so it can be assigned to a range safely.
'
Public Function VLookupLike_RowSlice2D( _
    ByVal ws As Worksheet, _
    ByVal startCol As Long, _
    ByVal endCol As Long, _
    ByVal firstDataRow As Long, _
    ByVal lastDataRow As Long, _
    ByVal identifier As String, _
    Optional ByVal caseSensitive As Boolean = False, _
    Optional ByVal trimMatchText As Boolean = True, _
    Optional ByVal handleMergedKeys As Boolean = True, _
    Optional ByRef matchedRow As Long = 0 _
) As Variant

    Dim outCols As Long
    Dim outArr() As Variant
    Dim r As Long, c As Long
    Dim idText As String, cellText As String
    Dim cmpMode As VbCompareMethod
    Dim keyCell As Range, ma As Range
    Dim topRow As Long, bottomRow As Long

    ' ---- Validation ----
    If ws Is Nothing Then Err.Raise vbObjectError + 810, , "Worksheet is required."
    If startCol < 1 Or endCol < 1 Then Err.Raise vbObjectError + 811, , "Column numbers must be >= 1."
    If endCol <= startCol Then Err.Raise vbObjectError + 812, , "endCol must be > startCol (there must be values to return)."
    If firstDataRow < 1 Or lastDataRow < firstDataRow Then Err.Raise vbObjectError + 813, , "Invalid row bounds."

    outCols = endCol - startCol                       ' exclude identifier column
    ReDim outArr(1 To 1, 1 To outCols)                ' prefill with empty strings
    For c = 1 To outCols
        outArr(1, c) = ""                             ' treat everything as string
    Next c

    idText = CStr(identifier)
    If trimMatchText Then idText = Trim$(idText)
    cmpMode = IIf(caseSensitive, vbBinaryCompare, vbTextCompare)

    ' ---- Search for first match in startCol ----
    r = firstDataRow
    Do While r <= lastDataRow

        Set keyCell = ws.Cells(r, startCol)

        If handleMergedKeys And keyCell.MergeCells Then
            Set ma = keyCell.MergeArea
            topRow = ma.row
            bottomRow = ma.row + ma.Rows.Count - 1

            ' Only evaluate once at the top row of the merged area
            If r = topRow Then
                cellText = CStr(IIf(IsEmpty(ma.Cells(1, 1).Value), "", ma.Cells(1, 1).Value))
                If trimMatchText Then cellText = Trim$(cellText)

                If StrComp(cellText, idText, cmpMode) = 0 Then
                    FillOutputRowStrings ws, topRow, startCol, endCol, outArr
                    matchedRow = topRow
                    VLookupLike_RowSlice2D = outArr
                    Exit Function
                End If
            End If

            ' Skip the rest of the merged block
            r = bottomRow + 1

        Else
            ' Non-merged (or not handling merged)
            cellText = CStr(IIf(IsEmpty(keyCell.Value), "", keyCell.Value))
            If trimMatchText Then cellText = Trim$(cellText)

            If StrComp(cellText, idText, cmpMode) = 0 Then
                FillOutputRowStrings ws, r, startCol, endCol, outArr
                matchedRow = r
                VLookupLike_RowSlice2D = outArr
                Exit Function
            End If

            r = r + 1
        End If
    Loop

    ' Not found: return the empty-initialized 1 x outCols 2-D array
    matchedRow = 0
    VLookupLike_RowSlice2D = outArr
End Function

' --- Helper ---

' Copies ws.Cells(row, startCol+1 .. endCol) into outArr(1, 1 .. endCol-startCol) as strings.
Private Sub FillOutputRowStrings( _
    ByVal ws As Worksheet, _
    ByVal row As Long, _
    ByVal startCol As Long, _
    ByVal endCol As Long, _
    ByRef outArr() As Variant _
)
    Dim c As Long, j As Long, v As Variant
    j = 1
    For c = startCol + 1 To endCol
        v = ws.Cells(row, c).Value
        outArr(1, j) = CStr(IIf(IsEmpty(v), "", v))
        j = j + 1
    Next c
End Sub

