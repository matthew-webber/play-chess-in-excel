VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Spawn_form 
   Caption         =   "Add a new piece to board"
   ClientHeight    =   2870
   ClientLeft      =   105
   ClientTop       =   455
   ClientWidth     =   3024
   OleObjectBlob   =   "Spawn_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Spawn_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub spawnbutton_Click()

Dim color As String
Dim rank As String
Dim colorrank As String

'color = colorbox.Value
'rank = rankbox.Value
colorrank = colorbox.Value & " " & rankbox.Value
'MsgBox colorrank & " spawned!", , "Leapin' Lizards!"
'MsgBox color & " " & rank
Range("K7").Font.Size = 14

If colorrank = "White King" Then
    Range("K7").Value = ChrW(&H2654)
ElseIf colorrank = "White Queen" Then
    Range("K7").Value = ChrW(&H2655)
ElseIf colorrank = "White Rook" Then
    Range("K7").Value = ChrW(&H2656)
ElseIf colorrank = "White Bishop" Then
    Range("K7").Value = ChrW(&H2657)
ElseIf colorrank = "White Knight" Then
    Range("K7").Value = ChrW(&H2658)
ElseIf colorrank = "White Pawn" Then
    Range("K7").Value = ChrW(&H2659)
ElseIf colorrank = "Black King" Then
    Range("K7").Value = ChrW(&H265A)
ElseIf colorrank = "Black Queen" Then
    Range("K7").Value = ChrW(&H265B)
ElseIf colorrank = "Black Rook" Then
    Range("K7").Value = ChrW(&H265C)
ElseIf colorrank = "Black Bishop" Then
    Range("K7").Value = ChrW(&H265D)
ElseIf colorrank = "Black Knight" Then
    Range("K7").Value = ChrW(&H265E)
ElseIf colorrank = "Black Pawn" Then
    Range("K7").Value = ChrW(&H265F)
    Range("K7").Font.Size = 9
End If

Unload Spawn_form


End Sub

Private Sub UserForm_Initialize()
    
    With colorbox
        .AddItem "White"
        .AddItem "Black"
    End With
    
    With rankbox
        .AddItem "Queen"
        .AddItem "King"
        .AddItem "Rook"
        .AddItem "Bishop"
        .AddItem "Knight"
        .AddItem "Pawn"
    End With
    
End Sub
