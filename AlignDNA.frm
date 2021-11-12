VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo - Global sequence alignment in VB6"
   ClientHeight    =   9945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   ScaleHeight     =   9945
   ScaleWidth      =   16020
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      Caption         =   "Options"
      Height          =   1575
      Left            =   2640
      TabIndex        =   25
      Top             =   1920
      Width           =   1815
      Begin VB.CheckBox SD 
         Caption         =   "Show diagonal"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         ToolTipText     =   "Show diagonal on Traceback matrix"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox PG 
         Caption         =   "Plot grid"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         ToolTipText     =   "Plot grid on Traceback matrix"
         Top             =   720
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox PTB 
         Caption         =   "Plot TraceBack"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         ToolTipText     =   "Plot path on the main matrix"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Presets"
      Height          =   4455
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   4335
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 16"
         Height          =   375
         Index           =   15
         Left            =   2280
         TabIndex        =   37
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 15"
         Height          =   375
         Index           =   14
         Left            =   2280
         TabIndex        =   36
         Top             =   3360
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 14"
         Height          =   375
         Index           =   13
         Left            =   2280
         TabIndex        =   35
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 13"
         Height          =   375
         Index           =   12
         Left            =   2280
         TabIndex        =   34
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 12"
         Height          =   375
         Index           =   11
         Left            =   2280
         TabIndex        =   33
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 11"
         Height          =   375
         Index           =   10
         Left            =   2280
         TabIndex        =   32
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 10"
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   29
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 9"
         Height          =   375
         Index           =   8
         Left            =   2280
         TabIndex        =   28
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 10"
         Height          =   375
         Index           =   7
         Left            =   2280
         TabIndex        =   27
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 1"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 2"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 3"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 4"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 5"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 6"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CommandButton Setting 
         Caption         =   "Experiment 7"
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   18
         Top             =   3360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Parameters"
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   2415
      Begin VB.TextBox Tgap 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Text            =   "-1"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox ma 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Text            =   "1"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox mma 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Text            =   "-1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Gap ="
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Match ="
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "MMatch ="
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sequences"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4335
      Begin VB.TextBox s1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Text            =   "GCATGCU"
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox s2 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Text            =   "GATTACA"
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "Sq 1 ="
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Sq 2 ="
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Graphic representation of the aligment matrix (colors represent values)"
      Height          =   4815
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   5535
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   240
         ScaleHeight     =   279
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   335
         TabIndex        =   4
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Traceback path deviation from diagonal"
      Height          =   4815
      Left            =   10320
      TabIndex        =   2
      Top             =   120
      Width           =   5535
      Begin VB.PictureBox Pic2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   240
         ScaleHeight     =   279
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   335
         TabIndex        =   38
         Top             =   360
         Width           =   5055
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderStyle     =   4  'Dash-Dot
            BorderWidth     =   3
            Visible         =   0   'False
            X1              =   0
            X2              =   368
            Y1              =   0
            Y2              =   304
         End
      End
   End
   Begin VB.CommandButton Align 
      Caption         =   "Align"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   3720
      Width           =   4095
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
      Height          =   3975
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   5040
      Width           =   11295
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   -9240
      Picture         =   "AlignDNA.frx":0000
      Top             =   9240
      Width           =   25290
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
' |          Update:  May 2021                                                    |
' |       Tested On:  WinXP, WinVista, Win7, Win8, Win10                          |
' |             Use:  sequence alignment                                          |
' |                                                                               |
' |                  _____________________________                                |
' |_________________/                             \_______________________________|
'

Dim Matrix() As Variant      'matrix of nucleotides
Dim MatrixTrace() As Variant 'matrix of path
Dim MatrixMaxVal As Variant  'matrix maximum value
Dim MatrixMinVal As Variant  'matrix minimum value

Private Sub Align_Click()

    '------ Filter and clear text OBJ -----------------------------------------
    s1.Text = UCase(s1.Text)
    s2.Text = UCase(s2.Text)
    
    s1.Text = Replace(s1.Text, Chr(13), "")
    s2.Text = Replace(s2.Text, Chr(13), "")
    
    s1.Text = Replace(s1.Text, " ", "")
    s2.Text = Replace(s2.Text, " ", "")
    
    result.Text = ""
    gap = Val(Tgap.Text)
    '--------------------------------------------------------------------------
    
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
    MatrixMaxVal = 0
    MatrixMinVal = 0
    
    For i = 2 To sq1 + 1 'Rows
    
        letter2 = Mid(s2.Text, i - 1, 1)
    
        For j = 2 To sq2 + 1 'Cols
        
            If i > 1 And j > 1 Then
        
                letter1 = Mid(s1.Text, j - 1, 1)
                A = Val(Matrix(i - 1, j - 1)) + S(letter1, letter2) '\
                B = Val(Matrix(i - 1, j)) + gap                     '-
                C = Val(Matrix(i, j - 1)) + gap                     '|
                
                Matrix(i, j) = MAX(A, B, C)
                
                If Val(Matrix(i, j)) > MatrixMaxVal Then MatrixMaxVal = Val(Matrix(i, j))
                If Val(Matrix(i, j)) < MatrixMinVal Then MatrixMinVal = Val(Matrix(i, j))
            
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
            MatrixTrace(i, j) = Matrix(i - 1, j) + gap
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
      
        If i <= 2 And j <= 2 Then MatrixTrace(i, j) = 0 ' sa puna un zero la (1,1)
      
    Loop
    
    '--------------------------------------------------------------------------
    If Len(AlignmentA) > Len(AlignmentB) Then k = Len(AlignmentA) Else k = Len(AlignmentB)
    
    cC = 0
    
    For i = 1 To k '+ 1
    
        l1 = Mid(AlignmentA, i, 1)
        l2 = Mid(AlignmentB, i, 1)
        
        If l1 = l2 Then
            AlignmentM = AlignmentM & "|"
            cC = cC + 1
        Else
            AlignmentM = AlignmentM & " "
        End If
        
    Next i
    '--------------------------------------------------------------------------
    
    result.Text = result.Text & "Show Alignment:" & vbCrLf & vbCrLf & AlignmentA & vbCrLf & AlignmentM & vbCrLf & AlignmentB & vbCrLf & vbCrLf
    
    result.Text = result.Text & "Matches = " & cC & vbCrLf & "Length = " & k & vbCrLf & vbCrLf
    
    result.Text = result.Text & "Similarity = " & Round(Val((100 / k) * cC), 2) & " %" & vbCrLf & vbCrLf
    
    If sq1 < 30 Or sq2 < 30 Then
        '------ Show TraceMatrix in Text OBJ --------------------------------------
        result.Text = result.Text & DrowMatrix(sq1 + 1, sq2 + 1, MatrixTrace, "Tracing back:")
        '--------------------------------------------------------------------------
        '------ Show Matrix in Text OBJ -------------------------------------------
        result.Text = result.Text & DrowMatrix(sq1 + 1, sq2 + 1, Matrix, "Show Matrix:")
        '--------------------------------------------------------------------------
    Else
        result.Text = result.Text & "Matrix too large to be shown in clear text !"
    End If
    
    
    
    result.Text = result.Text & ct & vbCrLf
    '--------------------------------------------------------------------------
    
    
    Call DrowColorMatrix(sq1 + 1, sq2 + 1, Matrix)
    
    Exit Sub
1:
    MsgBox "Error!" & vbCrLf & "One of the parameters is inadequate for the given conditions !"

End Sub


Function S(a1, a2) As Variant
    If a1 = a2 Then S = Val(ma.Text) Else S = Val(mma.Text)
End Function


Public Function MAX(ByVal ma, ByVal mb, ByVal mc)
    ma = IIf(ma > mb, ma, mb)
    ma = IIf(ma > mc, ma, mc)
    MAX = ma
End Function


Function DrowColorMatrix(ib, jb, ByVal M As Variant)
    '--------------------------------------------------------------------------
    Pic1.Cls
    Pic2.Cls
     
    'Row = (picOBJ.ScaleWidth / (jb + 1))
    'Col = (picOBJ.ScaleHeight / (ib + 1))
    
    Row = (Pic1.ScaleWidth / jb)
    Col = (Pic1.ScaleHeight / ib)
    
    Maxim = MatrixMaxVal
    Minim = Abs(MatrixMinVal) 'Abs() transforma nr negativ in pozitiv
    
    
    If Maxim <> 0 Then Culoare1 = Int(255 / Maxim)
    If Minim <> 0 Then Culoare2 = Int(255 / Minim)
    
    
    For i = 0 To jb 'Rows
    
        For j = 0 To ib 'cols
        
            'h = M(i + 1, j + 1)
            h = M(j + 1, i + 1)
            
            If h > 0 Then r = Culoare1 * h
            If h < 0 Then g = Culoare2 * Abs(h)
            
            Pic1.Line (Row * i, Col * j)-(Row * (i + 1), Col * (j + 1)), RGB(r, g, 55), BF
            
            If MatrixTrace(j + 1, i + 1) <> "" Then
                If PTB.Value = 1 Then Pic1.Line (Row * i, Col * j)-(Row * (i + 1), Col * (j + 1)), RGB(255, 255, 255), BF
                Pic2.Line (Row * i, Col * j)-(Row * (i + 1), Col * (j + 1)), RGB(200, 45, 45), BF
                
                If PG.Value = 1 Then
                    Pic2.Line (Row * i, 0)-(Row * i, Pic1.ScaleHeight), RGB(45, 45, 45), B
                    Pic2.Line (0, Col * j)-(Pic1.ScaleWidth, Col * j), RGB(45, 45, 45), B
                End If
            End If
            
        Next j
    
    Next i
    '--------------------------------------------------------------------------
End Function


Function DrowMatrix(ib, jb, ByVal M As Variant, ByVal msg As String) As String

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
    DrowMatrix = msg & " M[" & Val(jb - 1) & "," & Val(ib - 1) & "]" & vbCrLf & ct & vbCrLf
    '--------------------------------------------------------------------------

End Function


Private Sub Form_Load()
    Call Setting_Click(14)
    Align_Click
End Sub

Private Sub s1_Change()
    Align_Click
End Sub

Private Sub s2_Change()
    Align_Click
End Sub

Private Sub SD_Click()
    If SD.Value = 1 Then Line1.Visible = True Else Line1.Visible = False
End Sub

Private Sub Setting_Click(Index As Integer)

    If Index = 0 Then
    
        Tgap.Text = 0.5
        ma.Text = 1
        mma.Text = 0
        
        s1.Text = "GAATTCAGTTA" '(sequence #1)
        s2.Text = "GGATCGA"     '(sequence #2)
        
    End If
    
    
    If Index = 1 Then
    
        Tgap.Text = 0
        ma.Text = 1
        mma.Text = 0
    
        s1.Text = "gcgcgtgcgcggaaggagccaaggtgaagttgtagcagtgtgtcagaagaggtgcgtggcaccatgctgtcccccgaggcggagcgggtgctgcggtacctggtcgaagtagaggagttg" '(sequence #1)
        s2.Text = "gacttgtggaacctacttcctgaaaataaccttctgtcctccgagctctccgcacccgtggatgacctgctcccgtacacagatgttgccacctggctggatgaatgtccgaatgaagcg" '(sequence #2)
        
    End If
    
    
    If Index = 2 Then
    
        Tgap.Text = 0
        ma.Text = 1
        mma.Text = 0
    
        s1.Text = "AGTGTTCCAG"  '(sequence #1)
        s2.Text = "AATCGTTACAG" '(sequence #2)
        
    End If
    
    
    If Index = 3 Then
    
        Tgap.Text = 0
        ma.Text = 1
        mma.Text = -1
    
        s1.Text = "TATATCTGGCTATCTACTG" '(sequence #1)
        s2.Text = "AGCGTGCAGCCAATAC"    '(sequence #2)

    End If
    
    
    If Index = 4 Then
    
        Tgap.Text = 0
        ma.Text = 1
        mma.Text = 0
    
        s1.Text = "ACCGTGAAGCCAATAC"           '(sequence #1)
        s2.Text = "TATAGTCTCGTATCTATCATCTACTA" '(sequence #2)

    End If
    
    
    If Index = 5 Then
    
        Tgap.Text = -1
        ma.Text = 1
        mma.Text = -2
    
        s1.Text = "ACCGTGAAGCCAATAC" '(sequence #1)
        s2.Text = "AGCGTGCAGCCAATAC" '(sequence #2)

    End If
    
    
    If Index = 6 Then
    
        Tgap.Text = -1
        ma.Text = 1
        mma.Text = -20
    
        s1.Text = "CGGGCTCTCTCACGTCTAC" '(sequence #1)
        s2.Text = "GCTCTCTCACGTCTAC"    '(sequence #2)

    End If
    
    
    If Index = 7 Then
    
        Tgap.Text = -1
        ma.Text = 1
        mma.Text = -1
    
        s1.Text = "GCATGCGACTAC" '(sequence #1)
        s2.Text = "GATTACAGTCAC" '(sequence #2)

    End If
    
    
    If Index = 8 Then
    
        Tgap.Text = -2
        ma.Text = 1
        mma.Text = -1
    
        s1.Text = "ACCGTGAAGCCAATAC"      '(sequence #1)
        s2.Text = "AGCGTGAAAAACAGCCAATAC" '(sequence #2)

    End If
    
    
    If Index = 9 Then
    
        Tgap.Text = -2
        ma.Text = 1
        mma.Text = -1
    
        s1.Text = "ACCGTGAAGCCAATAC"        '(sequence #1)
        s2.Text = "AGCGTGAAAAAAAAGGCCAATAC" '(sequence #2)

    End If
    
    
    If Index = 10 Then
    
        Tgap.Text = -1
        ma.Text = 1
        mma.Text = -1
    
        s1.Text = "AAAAAAAAAAAAAAAAAAAAAAAAA" '(sequence #1)
        s2.Text = "AAAAAAAAAAAAAAAAAAAAAAAAA" '(sequence #2)

    End If
    
    
    If Index = 11 Then
    
        Tgap.Text = -1
        ma.Text = 1
        mma.Text = -1
    
        s1.Text = "AAAAAAAAAAAAAAAGGGAAAAAAAAAA" '(sequence #1)
        s2.Text = "AAATTTAAAAAAAAAAAAAAAAAAA"    '(sequence #2)

    End If
    
    
    If Index = 12 Then
    
        Tgap.Text = -1
        ma.Text = 1
        mma.Text = -1
    
        s1.Text = "AAAAACCCCAAAAAAAAAAATTTTTTTTTTTTTAAAA" '(sequence #1)
        s2.Text = "AAACCCCAATTTTTTTAAAAAAAAAAAAAAAAAA"    '(sequence #2)

    End If
    
    
    If Index = 13 Then
    
        Tgap.Text = -1
        ma.Text = 1
        mma.Text = -1
    
        s1.Text = "AAAAAAAAAAAAAAAAAAAAAAAAA"           '(sequence #1)
        s2.Text = "AAACCCCAATTTTGTTTAAAAAAAAAAAAAAAAAA" '(sequence #2)

    End If
    
    
    If Index = 14 Then
    
        Tgap.Text = -1
        ma.Text = 1
        mma.Text = -1
    
        s1.Text = "AAAATTTTAAAACCCCCAAAAAAAAAAA" '(sequence #1)
        s2.Text = "AAATTTTAAAAAAAAAACCCCCCAAAAA"   '(sequence #2)

    End If
    
    
    If Index = 15 Then
    
        Tgap.Text = -1
        ma.Text = 1
        mma.Text = -1
    
        s1.Text = "AAGGTTTGTACTGTAAGTAAGGTAAATTCGGTACGGCGGGTGAGGAAGGTGAATAAGGTAGGA" '(sequence #1)
        s2.Text = "AGAGTAATTCTTGTAAGTCCGGTACGCCAGGTGAAAGATGTACATATGGTTAGAGTGGTTTAA" '(sequence #2)

    End If


End Sub


