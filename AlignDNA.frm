VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo - Global sequence alignment in VB6"
   ClientHeight    =   12360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   ScaleHeight     =   12360
   ScaleWidth      =   12285
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll_mma 
      Height          =   255
      Left            =   7920
      Max             =   20
      Min             =   -20
      TabIndex        =   13
      Top             =   960
      Value           =   -1
      Width           =   3015
   End
   Begin VB.HScrollBar HScroll_ma 
      Height          =   255
      Left            =   7920
      Max             =   20
      Min             =   -20
      TabIndex        =   12
      Top             =   600
      Value           =   6
      Width           =   3015
   End
   Begin VB.HScrollBar HScroll_Tgap 
      Height          =   255
      Left            =   7920
      Max             =   20
      Min             =   -20
      TabIndex        =   11
      Top             =   240
      Value           =   -6
      Width           =   3015
   End
   Begin VB.TextBox mma 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   11040
      TabIndex        =   7
      Text            =   "-1"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox ma 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   11040
      TabIndex        =   5
      Text            =   "6"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Tgap 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   11040
      TabIndex        =   3
      Text            =   "-6"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox result 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10095
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1440
      Width           =   11895
   End
   Begin VB.TextBox s2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Text            =   "GTTATTACCATTGA"
      Top             =   840
      Width           =   6015
   End
   Begin VB.TextBox s1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Text            =   "GTTTGCATGCTTG"
      Top             =   240
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   -12960
      Picture         =   "AlignDNA.frx":0000
      Top             =   11640
      Width           =   25290
   End
   Begin VB.Label Label5 
      Caption         =   "Sq 2 ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Sq 1 ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "MMatch ="
      Height          =   255
      Left            =   7080
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Match ="
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Gap ="
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ________________________________                          ____________________
'  /    Global sequence alignment   \________________________/       v1.00        |
' |                                                                               |
' |            Name:  Demo - Global sequence alignment V1.0                       |
' |        Category:  Open source software                                        |
' |          Author:  Paul A. Gagniuc                                             |
' |            Book:  Algorithms in Bioinformatics: Theory and Implementation     |
' |           Email:  paul_gagniuc@acad.ro                                        |
' |  ____________________________________________________________________________ |
' |                                                                               |
' |    Date Created:  January 2014                                                |
' |          Update:  January 2021                                                |
' |       Tested On:  WinXP, WinVista, Win7, Win8, Win10                          |
' |             Use:  sequence alignment                                          |
' |                                                                               |
' |                  _____________________________                                |
' |_________________/                             \_______________________________|
'

Dim Matrix() As Variant      'matrix of nucleotides
Dim MatrixTrace() As Variant 'matrix of path


Private Sub Align_Click()
    
    result.Text = ""
    gap = Val(Tgap.Text)
    
    '------ Matrix(n,n) umplut cu caractere -----------------------------------
    sq1 = Len(s1.Text)
    sq2 = Len(s2.Text)
    
    sp = 2
    
    ReDim Matrix(0 To sq1 + sp, 0 To sq2 + sp) As Variant
    ReDim MatrixTrace(0 To sq1 + sp, 0 To sq2 + sp) As Variant
    
    Matrix(0, 0) = " "
    Matrix(1, 0) = " "
    Matrix(0, 1) = " "
    
    MatrixTrace(0, 0) = " "
    MatrixTrace(1, 0) = " "
    MatrixTrace(0, 1) = " "
    
    '------ Matrix(i,0) umplut cu s2 pe Row -----------------------------------
    For i = 2 To sq1 + sp 'Rows
        letter2 = Mid(s1.Text, i - 1, 1)
        Matrix(i, 0) = letter2
        MatrixTrace(i, 0) = letter2
    Next i
    '--------------------------------------------------------------------------
    
    '------ Matrix(0,j) umplut cu s1 pe Col -----------------------------------
    For j = 2 To sq2 + sp 'cols
        letter1 = Mid(s2.Text, j - 1, 1)
        Matrix(0, j) = letter1
        MatrixTrace(0, j) = letter1
    Next j
    '--------------------------------------------------------------------------
    
    '---- Fill Matrix Zero ----------------------------------------------------
    For i = 1 To sq1 + 1
        For j = 1 To sq2 + 1
            Matrix(i, j) = 0
        Next j
    Next i
    '--------------------------------------------------------------------------
    
    '---- Fill Matrix Extremis ------------------------------------------------
    For i = 1 To sq1
        Matrix(i + 1, 1) = Matrix(i, 1) - 1
    Next i
    '--------------------------------------------------------------------------
    
    '---- Fill Matrix Extremis ------------------------------------------------
    For j = 1 To sq2
        Matrix(1, j + 1) = Matrix(1, j) - 1
    Next j
    '--------------------------------------------------------------------------
    
    '---- Fill Matrix ---------------------------------------------------------
    For i = 2 To sq1 + 1 'Rows
    
        letter2 = Mid(s2.Text, i - 1, 1)
    
        For j = 2 To sq2 + 1 'Cols
        
            If i > 1 And j > 1 Then
        
                letter1 = Mid(s1.Text, j - 1, 1)
                A = Val(Matrix(i - 1, j - 1)) + S(letter1, letter2) '\
                B = Val(Matrix(i - 1, j)) + gap                     '-
                C = Val(Matrix(i, j - 1)) + gap                     '|
                
                Matrix(i, j) = VMAX(A, B, C)
            
            End If
    
        Next j
    
    Next i
    '--------------------------------------------------------------------------
    
    '---- Read Path Matrix ---------------------------------------------------------
    On Error GoTo 1
    'On Error Resume Next
        
    AlignmentA = ""
    AlignmentM = ""
    AlignmentB = ""
    i = sq1 + 1 'Rows
    j = sq2 + 1 'Cols
    
    Do While (i >= 2 Or j >= 2)
        
        Ai = Matrix(i, 0) 'Cols
        Bj = Matrix(0, j) 'Rows
    
        If (i >= 2 And j >= 2 And Matrix(i, j) = Val(Matrix(i - 1, j - 1)) + S(Ai, Bj)) Then
            MatrixTrace(i, j) = Matrix(i - 1, j - 1) + S(Ai, Bj)
            AlignmentA = Ai + AlignmentA
            AlignmentB = Bj + AlignmentB
            i = i - 1
            j = j - 1
        
        ElseIf (i >= 2 And Matrix(i, j) = Val(Matrix(i - 1, j)) + gap) Then
            MatrixTrace(i, j) = Val(Matrix(i - 1, j)) + gap
            AlignmentA = Ai + AlignmentA
            AlignmentB = "-" + AlignmentB
            i = i - 1
        ElseIf (j >= 2) Then
            MatrixTrace(i, j) = Val(Matrix(i, j - 1)) + gap
            AlignmentA = "-" + AlignmentA
            AlignmentB = Bj + AlignmentB
            j = j - 1
        Else
            MatrixTrace(i, j) = Val(Matrix(i - 1, j)) + gap
            AlignmentA = Ai + AlignmentA
            AlignmentB = "-" + AlignmentB
            i = i - 1
        End If
      
        If i <= 2 And j <= 2 Then MatrixTrace(i, j) = 0
      
    Loop
    
    '--------------------------------------------------------------------------
    If sq1 > sq2 Then k = sq1 Else k = sq2
    
    For i = 1 To k '- 1
    
        l1 = Mid(AlignmentA, i, 1)
        l2 = Mid(AlignmentB, i, 1)
        If l1 = l2 Then AlignmentM = AlignmentM & "|" Else AlignmentM = AlignmentM & " "
    
    Next i
    '--------------------------------------------------------------------------
    
    result.Text = result.Text & "Show alignment:" & vbCrLf & vbCrLf & AlignmentA & vbCrLf & AlignmentM & vbCrLf & AlignmentB & vbCrLf & vbCrLf
    
    '------ Show TraceMatrix in Text OBJ --------------------------------------
    Call DrowMatrix(sq1 + 1, sq2 + 1, MatrixTrace, "Plot Trace:")
    '--------------------------------------------------------------------------
    
    '------ Show Matrix in Text OBJ -------------------------------------------
    Call DrowMatrix(sq1 + 1, sq2 + 1, Matrix, "Plot Matrix:")
    '--------------------------------------------------------------------------
    result.Text = result.Text & ct & vbCrLf
    '--------------------------------------------------------------------------
    Exit Sub
1:
    MsgBox "Error!" & vbCrLf & "One of the parameters is inadequate for the given conditions !"
    
End Sub


Function S(a1, a2) As Variant
    If a1 = a2 Then S = Val(ma.Text) Else S = Val(mma.Text)
End Function

Function VMAX(a1, a2, a3) As Variant
    If a1 > a2 Then ctmp = a1 Else ctmp = a2
    If ctmp > a3 Then VMAX = ctmp Else VMAX = a3
End Function


Function DrowMatrix(ib, jb, ByVal M As Variant, ByVal msg As String) As Variant

    '------ Show Matrix in Text OBJ -------------------------------------------
    
    For i = 0 To jb 'Cols
        x = x & "_____"
    Next i
    
    ct = ct & vbCrLf & " " & x & vbCrLf
    
    For i = 0 To ib 'Rows
    
        For j = 0 To jb 'cols
        
            If Len(M(i, j)) = 0 Then u = "|    "
            If Len(M(i, j)) = 1 Then u = "|   "
            If Len(M(i, j)) = 2 Then u = "|  "
            If Len(M(i, j)) = 3 Then u = "| "
            If Len(M(i, j)) = 4 Then u = "|"
            
            If j = jb Then o = "|" Else o = ""
            ct = ct & u & M(i, j) & o
            If i = 0 Then y = y & "|____" & o
             
        Next j
    
    ct = ct & vbCrLf & y & vbCrLf
    
    Next i
    '--------------------------------------------------------------------------
    result.Text = result.Text & msg & vbCrLf & ct & vbCrLf
    '--------------------------------------------------------------------------

End Function

Private Sub Form_Load()
    Align_Click
End Sub

Private Sub HScroll_ma_Change()
    ma.Text = HScroll_ma.Value
    Align_Click
End Sub

Private Sub HScroll_ma_Scroll()
    ma.Text = HScroll_ma.Value
    Align_Click
End Sub

Private Sub HScroll_mma_Change()
    mma.Text = HScroll_mma.Value
    Align_Click
End Sub

Private Sub HScroll_mma_Scroll()
    mma.Text = HScroll_mma.Value
    Align_Click
End Sub

Private Sub HScroll_Tgap_Change()
    Tgap.Text = HScroll_Tgap.Value
    Align_Click
End Sub

Private Sub HScroll_Tgap_Scroll()
    Tgap.Text = HScroll_Tgap.Value
    Align_Click
End Sub

Private Sub s1_Change()
    Align_Click
End Sub

Private Sub s2_Change()
    Align_Click
End Sub
