Attribute VB_Name = "Module1"
Function mainmatch()
Dim matchcolor As Integer
Dim foundtemp As Integer

highmatch = 0
lastrow = Sheets("Acq-Div List").Range("A" & Rows.Count).End(xlUp).Row

For Each cell In Sheets("Acq-Div List").Range("A2:A" & lastrow)
    foundtemp = 1
    For i = 1 To 4
        If Not IsError(Application.Match(Sheets("Acq-Div List").Cells(cell.Row, i).Value, Sheets("Lithia List").Columns(i), 0)) Then
            foundtemp = Application.Match(Sheets("Acq-Div List").Cells(cell.Row, i).Value, Sheets("Lithia List").Columns(i), 0)
            found = found & foundtemp & ","
        Else
            found = found & 0 & ","
        End If
    Next i
    
        If Sheets("Acq-Div List").Cells(cell.Row, 5).Value = Sheets("Lithia List").Cells(foundtemp, 5).Value Then
            found = found & foundtemp & ","
        Else
            found = found & 0 & ","
        End If
    
    
    
    foundmtx = Split(found, ",")
    
    For y = 0 To 4
        If Not foundmtx(y) = 0 Then
            matched = 0
            
            For x = 0 To 4
                If foundmtx(x) = foundmtx(y) Then
                    matched = matched + 1
                    
                End If
            Next x
            
            If matched > highmatch Then
                highmatch = matched
                matchedrow = foundmtx(y)
            End If
            
            
        End If
    Next y
    
    Select Case highmatch
        Case 0
            matchcolor = 2
        Case 1, 2
            matchcolor = 3
            
        Case 3
            matchcolor = 45
            
        Case 4
            matchcolor = 6
        Case 5
            matchcolor = 4
    End Select
    
    Sheets("match").Range("R" & cell.Row).Value = highmatch
    Sheets("match").Range("R" & cell.Row).Interior.ColorIndex = matchcolor
    For j = 0 To 4
        If foundmtx(j) = matchedrow Then
            Sheets("Match").Cells(cell.Row, j + 1).Interior.ColorIndex = matchcolor
        End If
    Next j
    
'    Sheets("match").Range("A" & cell.Row).Value = cell.Value
'    Sheets("match").Range("B" & cell.Row).Value = highmatch
    highmatch = 0
    found = ""
If cell.Row = lastrow Then
    MsgBox ("Done")
End If

Next cell
mainmatch = 1
End Function
