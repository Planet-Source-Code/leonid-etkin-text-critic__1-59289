VERSION 5.00
Begin VB.Form FrmCritic 
   Caption         =   "TextCritic (Tedious, line-by-line, hidden meaning seaking, essay, novel, poem, and other text analyzer)"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCritique 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   960
      Width           =   5415
   End
End
Attribute VB_Name = "FrmCritic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VocWord(1 To 11) As Word

Public Sub Load_Vocabulary()
Dim strpump, vocpump, setinput
    Open App.Path & "\" & "CriticsVocabulary.txt" For Input As #1
    'Start reading info
    Do While Not EOF(1)
        DoEvents
        Input #1, strpump
        'If retrieved number (the beginning of the next set)
        'Start reading the vocabulary
        If IsNumeric(strpump) And strpump <> Empty Then
                setinput = 0
                'First input to prevent one blank input
                'And one extra input at the end
                Input #1, vocpump
            Do While vocpump <> "..."
                ReDim Preserve VocWord(strpump).wSet(setinput)
                
                VocWord(strpump).Size = VocWord(strpump).Size + 1
                
                VocWord(strpump).wSet(setinput) = vocpump
                setinput = setinput + 1
                Input #1, vocpump
            Loop
        End If
            'Word (strpump)
    Loop
    Close #1 'Close File
End Sub

Private Sub Form_DblClick()
'    Test_Vocabulary
End Sub

Private Sub Form_Load()
    Load_Vocabulary
End Sub

Public Sub Critisize(WritBrd As Object)
Dim nxtrnd, x, tway, startx 'next random word

txtCritique.Text = txtCritique.Text
Randomize
tway = Int(Rnd(1) * 2)
If tway = 1 Then
    startx = 1
Else
    startx = 2
End If


For x = startx To 3
    Randomize
    nxtrnd = Int(Rnd(1) * VocWord(x).Size)
    WritBrd = WritBrd & " " & VocWord(x).wSet(nxtrnd)
Next

    nxtrnd = Int(Rnd(1) * VocWord(4).Size)
    WritBrd = WritBrd & " " & VocWord(4).wSet(nxtrnd)
If nxtrnd = 0 Then
    nxtrnd = Int(Rnd(1) * VocWord(5).Size)
    WritBrd = WritBrd & " " & VocWord(5).wSet(nxtrnd)
Else
    nxtrnd = Int(Rnd(1) * VocWord(6).Size)
    WritBrd = WritBrd & " " & VocWord(6).wSet(nxtrnd)
End If

For x = 7 To 9
    Randomize
    nxtrnd = Int(Rnd(1) * VocWord(x).Size)
    WritBrd = WritBrd & " " & VocWord(x).wSet(nxtrnd)
Next
    nxtrnd = Int(Rnd(1) * VocWord(4).Size)
    WritBrd = WritBrd & " " & "the" & " " & VocWord(4).wSet(nxtrnd)

    tway = Int(Rnd(1) * 1)
    If tway = 0 Then
        nxtrnd = Int(Rnd(1) * VocWord(5).Size)
        WritBrd = WritBrd & " " & VocWord(5).wSet(nxtrnd)
    ElseIf tway = 1 Then
        nxtrnd = Int(Rnd(1) * VocWord(6).Size)
        WritBrd = WritBrd & " " & VocWord(6).wSet(nxtrnd)
    End If
        nxtrnd = Int(Rnd(1) * VocWord(10).Size)
        WritBrd = WritBrd & " " & VocWord(10).wSet(nxtrnd)
        
        nxtrnd = Int(Rnd(1) * VocWord(3).Size)
        WritBrd = WritBrd & " " & "the" & " " & VocWord(3).wSet(nxtrnd)
        
        nxtrnd = Int(Rnd(1) * VocWord(6).Size)
    WritBrd = WritBrd & " " & VocWord(6).wSet(nxtrnd)
    
    nxtrnd = Int(Rnd(1) * VocWord(11).Size)
        WritBrd = WritBrd & " " & VocWord(11).wSet(nxtrnd)
        
        WritBrd = WritBrd & SkipLine(2)
End Sub

Private Sub Test_Vocabulary()
On Error Resume Next
Dim x, y
    For x = 1 To 11
        DoEvents
        For y = 1 To VocWord(x).Size
            DoEvents
            Debug.Print VocWord(x).wSet(y)
        Next
        Debug.Print
    Next
End Sub

Private Sub txtCritique_DblClick()
  Critisize txtCritique
End Sub

Public Function SkipLine(count As Integer) As String
Dim x
For x = 1 To count
    SkipLine = SkipLine & Chr(13) & Chr(10)
Next
End Function
