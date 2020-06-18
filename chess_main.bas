Attribute VB_Name = "chess_macro"
Sub setup()

thisthat = MsgBox("Welcome to Chess for Excel!" _
& vbNewLine _
& vbNewLine _
& "To move a piece, click the piece you would like to move and Ctrl+Click the destination square." _
& "  Then click the 'Move Piece' button.  If you capture a piece by accident, use the 'Spawn Piece'" _
& " button to create a new one and then just copy+paste the replacement piece wherever you need it." _
& vbNewLine _
& vbNewLine _
& "Use the 'New Game' button to start a new game as either black or white." _
& "  After setup, if you need to swap to black, use this button and just start a new game." _
& vbNewLine _
& vbNewLine _
& "Enjoy!!", vbOKOnly, "Chess for Excel")

'===CREATE CHESS SHEET===
Dim sheet As Worksheet
Set sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
sheet.Name = "Chess!"
sheet.Activate

ActiveWindow.DisplayHeadings = False 'Hide row/column headers

new_game_setup
color_board
button_setup

End Sub

Sub color_board()

'Fill colors of chessboard, add borders to chessboard

   Const NB_CELLS As Integer = 8 '8x8 chessboard of cells
   Dim offset_row As Integer, offset_col As Integer ' => adding 2 variables
    
   Range("B2:I9").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
   Range("B2").Select
   
   'Shift (rows) starting from the first cell = the row number of the active cell - 1
   offset_row = ActiveCell.Row - 1
   'Shift (columns) starting from the first cell = the column number of the active cell - 1
   offset_col = ActiveCell.Column - 1
   
   For r = 1 To NB_CELLS 'Row number
   
        For c = 1 To NB_CELLS 'Column number
       
            If (r + c) Mod 2 = 0 Then
            'Cells(row number + number of rows to shift, column number + number of columns to shift)
                With Cells(r + offset_row, c + offset_col).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.149998474074526
                    .PatternTintAndShade = 0
                End With
            Else
                With Cells(r + offset_row, c + offset_col).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        Next
   Next
   
End Sub
Sub show_form()

    Spawn_form.Show
    
End Sub
Sub button_setup()

'===INSERT BUTTONS===

For i = 1 To 3
    If i = 1 Then
        Set btn = ActiveSheet.Buttons.Add(200, 75, 67.2, 22.8)
        With btn
            .Name = "movepiece"
            .Caption = "Move Piece"
            .OnAction = "piece_mover"
        End With
    ElseIf i = 2 Then
        Set btn = ActiveSheet.Buttons.Add(275, 75, 67.2, 22.8)
        With btn
            .Name = "newgame"
            .Caption = "New Game"
            .OnAction = "new_game"
        End With
    ElseIf i = 3 Then
        Set btn = ActiveSheet.Buttons.Add(350, 75, 67.2, 22.8)
        With btn
            .Name = "spawnpiece"
            .Caption = "Spawn Piece"
            .OnAction = "show_form"
        End With
    End If
Next i

''===MAKE THE SPAWN PIECE BUTTON (WHITE)===
'
'' Make the button.
''Dim obj As OLEObject
'Set obj = ActiveSheet.OLEObjects.Add( _
'    ClassType:="Forms.CommandButton.1", _
'    Link:=False, DisplayAsIcon:=False, _
'    Left:=350, Top:=75, _
'    Width:=67.2, Height:=22.8)
'' Get the button from the OLEObject.
'Dim cmd As CommandButton
'Set cmd = obj.Object
'
'' Set the button's name and picture.
'obj.Name = "CommandButton1"
'cmd.Caption = "Spawn Piece"
'
'' Select it to make it visible.
'obj.Select

End Sub

Sub new_game()

If MsgBox("Begin a new game?", vbOKCancel, "New game") = vbCancel Then
    Exit Sub
Else
        
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'===DELETE BOARD & BUTTONS===

    Range("A1:K11").Select
    Selection.Clear
'    ActiveSheet.Shapes.SelectAll
'    Selection.Delete
    
If MsgBox("If you are playing WHITE, click OK.  If you are playing BLACK, click cancel.", vbOKCancel, "Which side are you on?") = vbCancel Then
    
    '===IF PLAYER IS BLACK===
    
    '===SPAWN PIECES===

    'WHITE PAWNS
    Range("B2").Value = ChrW(&H2656)
    Range("C2").Value = ChrW(&H2658)
    Range("D2").Value = ChrW(&H2657)
    Range("E2").Value = ChrW(&H2655)
    Range("F2").Value = ChrW(&H2654)
    Range("G2").Value = ChrW(&H2657)
    Range("H2").Value = ChrW(&H2658)
    Range("I2").Value = ChrW(&H2656)
    
    'BLACK ROYALTY
    Range("B9").Value = ChrW(&H265C)
    Range("C9").Value = ChrW(&H265E)
    Range("D9").Value = ChrW(&H265D)
    Range("E9").Value = ChrW(&H265B)
    Range("F9").Value = ChrW(&H265A)
    Range("G9").Value = ChrW(&H265D)
    Range("H9").Value = ChrW(&H265E)
    Range("I9").Value = ChrW(&H265C)
    
    'WHITE PAWNS
    Range("B3").Value = ChrW(&H2659)
    Range("C3").Value = ChrW(&H2659)
    Range("D3").Value = ChrW(&H2659)
    Range("E3").Value = ChrW(&H2659)
    Range("F3").Value = ChrW(&H2659)
    Range("G3").Value = ChrW(&H2659)
    Range("H3").Value = ChrW(&H2659)
    Range("I3").Value = ChrW(&H2659)
    
    'BLACK PAWNS
    Range("B8").Value = ChrW(&H265F)
    Range("C8").Value = ChrW(&H265F)
    Range("D8").Value = ChrW(&H265F)
    Range("E8").Value = ChrW(&H265F)
    Range("F8").Value = ChrW(&H265F)
    Range("G8").Value = ChrW(&H265F)
    Range("H8").Value = ChrW(&H265F)
    Range("I8").Value = ChrW(&H265F)
    
    'Fill colors of chessboard, add borders to chessboard

   color_board
   
   '===ALGEBRAIC NOTATION HEADERS===
    
    Range("A2").Select
    
    i = 8
    y = 9
    
    Do While i > 0
        Cells(y, 1).Value = i
        y = y - 1
        i = i - 1
    Loop
    
    Range("B10").Value = "H"
    Range("C10").Value = "G"
    Range("D10").Value = "F"
    Range("E10").Value = "E"
    Range("F10").Value = "D"
    Range("G10").Value = "C"
    Range("H10").Value = "B"
    Range("I10").Value = "A"
    
    '===ROW/COLUMN HEIGHT + ALIGNMENT===

    Columns("A:I").ColumnWidth = 3.14
    Rows("1:10").RowHeight = 20.25
    With Range("A2:I10")
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With Range("B8:I8")
        .Font.Size = 9  ' Shrink black pawns bc Microsoft renders black pawns as large emojis instead of Unicode
    End With


Else

    '==IF PLAYER IS WHITE==
    
    '===SPAWN PIECES===

    'BLACK ROYALTY
    Range("B2").Value = ChrW(&H265C)
    Range("C2").Value = ChrW(&H265E)
    Range("D2").Value = ChrW(&H265D)
    Range("E2").Value = ChrW(&H265B)
    Range("F2").Value = ChrW(&H265A)
    Range("G2").Value = ChrW(&H265D)
    Range("H2").Value = ChrW(&H265E)
    Range("I2").Value = ChrW(&H265C)

    'BLACK PAWNS
    Range("B3").Value = ChrW(&H265F)
    Range("C3").Value = ChrW(&H265F)
    Range("D3").Value = ChrW(&H265F)
    Range("E3").Value = ChrW(&H265F)
    Range("F3").Value = ChrW(&H265F)
    Range("G3").Value = ChrW(&H265F)
    Range("H3").Value = ChrW(&H265F)
    Range("I3").Value = ChrW(&H265F)

    'WHITE ROYALTY
    Range("B9").Value = ChrW(&H2656)
    Range("C9").Value = ChrW(&H2658)
    Range("D9").Value = ChrW(&H2657)
    Range("E9").Value = ChrW(&H2655)
    Range("F9").Value = ChrW(&H2654)
    Range("G9").Value = ChrW(&H2657)
    Range("H9").Value = ChrW(&H2658)
    Range("I9").Value = ChrW(&H2656)

    'WHITE PAWNS
    Range("B8").Value = ChrW(&H2659)
    Range("C8").Value = ChrW(&H2659)
    Range("D8").Value = ChrW(&H2659)
    Range("E8").Value = ChrW(&H2659)
    Range("F8").Value = ChrW(&H2659)
    Range("G8").Value = ChrW(&H2659)
    Range("H8").Value = ChrW(&H2659)
    Range("I8").Value = ChrW(&H2659)
      
    color_board

    '===ROW/COLUMN HEIGHT===

    Columns("A:I").ColumnWidth = 3.14
    Rows("1:10").RowHeight = 20.25
    With Range("A1:I10")
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With Range("B3:I3")
        .Font.Size = 9  ' Shrink black pawns bc Microsoft renders black pawns as large emojis instead of Unicode
    End With
    
    ' ===INSERT NEW ROW/COLUMN HEADERS===

    Range("A2").Select

    i = 8
    y = 2

    Do While i > 0
        Cells(y, 1).Value = i
        y = y + 1
        i = i - 1
    Loop

    Range("B10").Value = "A"
    Range("C10").Value = "B"
    Range("D10").Value = "C"
    Range("E10").Value = "D"
    Range("F10").Value = "E"
    Range("G10").Value = "F"
    Range("H10").Value = "G"
    Range("I10").Value = "H"
    
End If
End If
    
Application.ScreenUpdating = True
Application.ScreenUpdating = True


End Sub

Sub new_game_setup()
        
Application.ScreenUpdating = False
Application.DisplayAlerts = False
 
 '===SPAWN PIECES===

 ActiveWindow.DisplayHeadings = False 'Hide row/column headers


 'BLACK ROYALTY
 Range("B2").Value = ChrW(&H265C)
 Range("C2").Value = ChrW(&H265E)
 Range("D2").Value = ChrW(&H265D)
 Range("E2").Value = ChrW(&H265B)
 Range("F2").Value = ChrW(&H265A)
 Range("G2").Value = ChrW(&H265D)
 Range("H2").Value = ChrW(&H265E)
 Range("I2").Value = ChrW(&H265C)

 'BLACK PAWNS
 Range("B3").Value = ChrW(&H265F)
 Range("C3").Value = ChrW(&H265F)
 Range("D3").Value = ChrW(&H265F)
 Range("E3").Value = ChrW(&H265F)
 Range("F3").Value = ChrW(&H265F)
 Range("G3").Value = ChrW(&H265F)
 Range("H3").Value = ChrW(&H265F)
 Range("I3").Value = ChrW(&H265F)

 'WHITE ROYALTY
 Range("B9").Value = ChrW(&H2656)
 Range("C9").Value = ChrW(&H2658)
 Range("D9").Value = ChrW(&H2657)
 Range("E9").Value = ChrW(&H2655)
 Range("F9").Value = ChrW(&H2654)
 Range("G9").Value = ChrW(&H2657)
 Range("H9").Value = ChrW(&H2658)
 Range("I9").Value = ChrW(&H2656)

 'WHITE PAWNS
 Range("B8").Value = ChrW(&H2659)
 Range("C8").Value = ChrW(&H2659)
 Range("D8").Value = ChrW(&H2659)
 Range("E8").Value = ChrW(&H2659)
 Range("F8").Value = ChrW(&H2659)
 Range("G8").Value = ChrW(&H2659)
 Range("H8").Value = ChrW(&H2659)
 Range("I8").Value = ChrW(&H2659)
   
Range("b2:i9").Select
 Selection.Borders(xlDiagonalDown).LineStyle = xlNone
 Selection.Borders(xlDiagonalUp).LineStyle = xlNone
 With Selection.Borders(xlEdgeLeft)
     .LineStyle = xlContinuous
     .ColorIndex = 0
     .TintAndShade = 0
     .Weight = xlThin
 End With
 With Selection.Borders(xlEdgeTop)
     .LineStyle = xlContinuous
     .ColorIndex = 0
     .TintAndShade = 0
     .Weight = xlThin
 End With
 With Selection.Borders(xlEdgeBottom)
     .LineStyle = xlContinuous
     .ColorIndex = 0
     .TintAndShade = 0
     .Weight = xlThin
 End With
 With Selection.Borders(xlEdgeRight)
     .LineStyle = xlContinuous
     .ColorIndex = 0
     .TintAndShade = 0
     .Weight = xlThin
 End With
 With Selection.Borders(xlInsideVertical)
     .LineStyle = xlContinuous
     .ColorIndex = 0
     .TintAndShade = 0
     .Weight = xlThin
 End With
 With Selection.Borders(xlInsideHorizontal)
     .LineStyle = xlContinuous
     .ColorIndex = 0
     .TintAndShade = 0
     .Weight = xlThin
 End With
 
Range("B2").Select

'Shift (rows) starting from the first cell = the row number of the active cell - 1
offset_row = ActiveCell.Row - 1
'Shift (columns) starting from the first cell = the column number of the active cell - 1
offset_col = ActiveCell.Column - 1

For r = 1 To NB_CELLS 'Row number

     For c = 1 To NB_CELLS 'Column number
    
         If (r + c) Mod 2 = 0 Then
         'Cells(row number + number of rows to shift, column number + number of columns to shift)
             With Cells(r + offset_row, c + offset_col).Interior
                 .Pattern = xlSolid
                 .PatternColorIndex = xlAutomatic
                 .ThemeColor = xlThemeColorDark1
                 .TintAndShade = -0.149998474074526
                 .PatternTintAndShade = 0
             End With
         Else
             With Cells(r + offset_row, c + offset_col).Interior
                 .Pattern = xlSolid
                 .PatternColorIndex = xlAutomatic
                 .ThemeColor = xlThemeColorDark1
                 .TintAndShade = 0
                 .PatternTintAndShade = 0
             End With
         End If
     Next
Next


 '===ROW/COLUMN HEIGHT===

 Columns("A:I").ColumnWidth = 3.14
 Rows("1:10").RowHeight = 20.25
 With Range("A1:I10")
     .Font.Size = 14
     .HorizontalAlignment = xlCenter
     .VerticalAlignment = xlCenter
 End With
 
 With Range("B3:I3")
     .Font.Size = 9  ' Shrink black pawns bc Microsoft renders black pawns as large emojis instead of Unicode
 End With
 
 '===INSERT NEW ROW/COLUMN HEADERS===

 Range("A2").Select

 i = 8
 y = 2

 Do While i > 0
     Cells(y, 1).Value = i
     y = y + 1
     i = i - 1
 Loop

 Range("B10").Value = "A"
 Range("C10").Value = "B"
 Range("D10").Value = "C"
 Range("E10").Value = "D"
 Range("F10").Value = "E"
 Range("G10").Value = "F"
 Range("H10").Value = "G"
 Range("I10").Value = "H"
    

Application.ScreenUpdating = True
Application.ScreenUpdating = True


End Sub


Sub piece_mover()

Dim cel As Range
Dim dest As Range
Dim src As Range

Set dest = Range(ActiveCell, ActiveCell)

For Each cel In Selection.Cells

Set src = Range(ActiveCell, ActiveCell)
cel.Activate

Next cel

dest.Value = src.Value
dest.Font.Size = src.Font.Size

src.Value = ""

End Sub


