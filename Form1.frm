VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Mine Sweeper"
   ClientHeight    =   3660
   ClientLeft      =   4455
   ClientTop       =   4080
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3660
   ScaleWidth      =   3090
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   60
      ScaleHeight     =   315
      ScaleWidth      =   1215
      TabIndex        =   39
      Top             =   1320
      Width           =   1215
      Begin VB.PictureBox picDisp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   420
         Picture         =   "Form1.frx":82A6A
         ScaleHeight     =   225
         ScaleWidth      =   120
         TabIndex        =   44
         Top             =   30
         Width           =   150
      End
      Begin VB.PictureBox picDisp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   480
         Picture         =   "Form1.frx":82C14
         ScaleHeight     =   225
         ScaleWidth      =   120
         TabIndex        =   43
         Top             =   30
         Width           =   150
      End
      Begin VB.PictureBox picDisp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   540
         Picture         =   "Form1.frx":82DBE
         ScaleHeight     =   225
         ScaleWidth      =   120
         TabIndex        =   42
         Top             =   30
         Width           =   150
      End
      Begin VB.PictureBox picDisp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   720
         Picture         =   "Form1.frx":82F68
         ScaleHeight     =   225
         ScaleWidth      =   120
         TabIndex        =   41
         Top             =   30
         Width           =   150
      End
      Begin VB.PictureBox picDisp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   900
         Picture         =   "Form1.frx":83112
         ScaleHeight     =   225
         ScaleWidth      =   120
         TabIndex        =   40
         Top             =   30
         Width           =   150
      End
   End
   Begin VB.PictureBox picDig 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   6600
      Picture         =   "Form1.frx":832BC
      ScaleHeight     =   225
      ScaleWidth      =   120
      TabIndex        =   38
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picDig 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   6420
      Picture         =   "Form1.frx":83466
      ScaleHeight     =   225
      ScaleWidth      =   120
      TabIndex        =   37
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picDig 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   6240
      Picture         =   "Form1.frx":83610
      ScaleHeight     =   225
      ScaleWidth      =   120
      TabIndex        =   36
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picDig 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   6060
      Picture         =   "Form1.frx":837BA
      ScaleHeight     =   225
      ScaleWidth      =   120
      TabIndex        =   35
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picDig 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   5880
      Picture         =   "Form1.frx":83964
      ScaleHeight     =   225
      ScaleWidth      =   120
      TabIndex        =   34
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picDig 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   5700
      Picture         =   "Form1.frx":83B0E
      ScaleHeight     =   225
      ScaleWidth      =   120
      TabIndex        =   33
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picDig 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   5520
      Picture         =   "Form1.frx":83CB8
      ScaleHeight     =   225
      ScaleWidth      =   120
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picDig 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   5340
      Picture         =   "Form1.frx":83E62
      ScaleHeight     =   225
      ScaleWidth      =   120
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picDig 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5160
      Picture         =   "Form1.frx":8400C
      ScaleHeight     =   225
      ScaleWidth      =   120
      TabIndex        =   30
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picDig 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4980
      Picture         =   "Form1.frx":841B6
      ScaleHeight     =   225
      ScaleWidth      =   120
      TabIndex        =   29
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   999
      Left            =   780
      Top             =   3060
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   60
      Picture         =   "Form1.frx":84360
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   28
      ToolTipText     =   "Hard"
      Top             =   660
      Width           =   285
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   60
      Picture         =   "Form1.frx":84716
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   27
      ToolTipText     =   "Nightmare"
      Top             =   960
      Width           =   285
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   60
      Picture         =   "Form1.frx":84ACC
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   26
      ToolTipText     =   "Normal"
      Top             =   360
      Width           =   285
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   60
      Picture         =   "Form1.frx":84E82
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   25
      ToolTipText     =   "Easy"
      Top             =   60
      Width           =   285
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reveal"
      Height          =   315
      Left            =   60
      TabIndex        =   24
      Top             =   2340
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   315
      Left            =   60
      TabIndex        =   22
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtX 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      TabIndex        =   18
      Text            =   "12"
      Top             =   60
      Width           =   555
   End
   Begin VB.TextBox txtY 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      TabIndex        =   17
      Text            =   "12"
      Top             =   360
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Game"
      Height          =   315
      Left            =   60
      TabIndex        =   16
      Top             =   1980
      Width           =   1215
   End
   Begin VB.TextBox txtMines 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      TabIndex        =   15
      Text            =   "18"
      Top             =   960
      Width           =   555
   End
   Begin VB.PictureBox picPlayArea 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   1380
      ScaleHeight     =   3465
      ScaleWidth      =   1605
      TabIndex        =   13
      Top             =   60
      Width           =   1635
      Begin VB.PictureBox picCell 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   60
         Picture         =   "Form1.frx":85238
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   14
         Top             =   60
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.PictureBox picNoBomb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   7680
      Picture         =   "Form1.frx":8554A
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   12
      Top             =   180
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picBomb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   7380
      Picture         =   "Form1.frx":8585C
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   11
      Top             =   180
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picMarked 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   7080
      Picture         =   "Form1.frx":85B6E
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   10
      Top             =   180
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   7
      Left            =   7380
      Picture         =   "Form1.frx":85E80
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   9
      Top             =   780
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   6
      Left            =   7080
      Picture         =   "Form1.frx":86192
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   780
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   5
      Left            =   6780
      Picture         =   "Form1.frx":864A4
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   7
      Top             =   780
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   4
      Left            =   6480
      Picture         =   "Form1.frx":867B6
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   780
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   7380
      Picture         =   "Form1.frx":86AC8
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   7080
      Picture         =   "Form1.frx":86DDA
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   6780
      Picture         =   "Form1.frx":870EC
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   6480
      Picture         =   "Form1.frx":873FE
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picFlat 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6780
      Picture         =   "Form1.frx":87710
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picBlank 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6480
      Picture         =   "Form1.frx":87A22
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   225
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   135
      Left            =   60
      TabIndex        =   23
      Top             =   1740
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   21
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   20
      Top             =   420
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mines"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   19
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************************
'** Minesweeper by Mike Toye, Jun 2001                **
'**                                                   **
'** Feel free to mess around with this code           **
'** If you make it faster, please let me know!        **
'**                                                   **
'**                              MADToye@aol.com      **
'*******************************************************

Dim iBombs() As Integer
Dim iMines As Integer
Dim itxtX As Integer, itxtY As Integer
Dim dTimer As Integer
Sub DefineGrid()
Dim x As Integer, y As Integer
Dim xCells As Integer
Dim xOffset As Long, yOffset As Long
Dim xCellGap As Long, yCellGap As Long
Dim xCellLen As Long, yCellLen As Long
Dim xPos As Long, yPos As Long

Dim tTime As Double
Dim lHeight As Long
    tTime = Timer
    xCells = CInt(txtX) * CInt(txtY)
    xOffset = 20
    yOffset = 20
    xCellGap = 0
    yCellGap = 0
    xCellLen = picBlank.Width
    yCellLen = picBlank.Height
    itxtX = CInt(txtX)
    itxtY = CInt(txtY)
    iMines = CInt(txtMines)
    picPlayArea.Width = (itxtX * (xCellLen + xCellGap)) + (2 * xOffset)
    picPlayArea.Height = (itxtY * (yCellLen + yCellGap)) + (2 * yOffset)
    Me.Width = picPlayArea.Left + picPlayArea.Width + 240
    lHeight = picPlayArea.Top + picPlayArea.Height + 500
    Me.Height = IIf(lHeight > (Command2.Top + Command2.Height + 500), lHeight, Command2.Top + Command2.Height + 500)
    PB.Value = 0
    PB.Max = xCells - 1
    PB.Visible = True
    DoEvents
    'right number of cells first!!
    If picCell.Count > xCells Then
        For x = xCells To picCell.Count - 1
            Unload picCell(x)
        Next x
    End If
    If xCells > picCell.Count Then
        For x = picCell.Count To xCells - 1
            Load picCell(x)
        Next x
    End If

    'display all as blank buttons
    xPos = xOffset
    yPos = yOffset
    For x = 0 To xCells - 1
        PB.Value = x
        
        picCell(x).Left = xPos
        picCell(x).Top = yPos
        picCell(x).Picture = picBlank.Picture
        picCell(x).AutoSize = True
        'picCell(X).BorderStyle = 0
        'picCell(X).Appearance = 0
        picCell(x).Visible = True
        If (x + 1) Mod itxtX = 0 Then
            yPos = yPos + yCellGap + yCellLen
            xPos = xOffset
        Else
            xPos = xPos + xCellGap + xCellLen
        End If

    Next x
    PB.Visible = False
    Me.Caption = picCell.Count & " cells in " & Timer - tTime & " seconds"

End Sub
Private Function CountTheBlanks() As Integer
Dim x As Integer, c As Integer
    c = 0
    For x = 0 To (itxtX * itxtY) - 1
        If picCell(x).Picture = picBlank.Picture Or _
            picCell(x).Picture = picMarked.Picture Then
            c = c + 1
        End If
    Next x
    CountTheBlanks = c
End Function
Private Sub Command1_Click()
    If Command1.Caption = "New Game" Then
        DefineGrid
        SetMines
        Command1.Caption = "Stop!"
        dTimer = 0
        picPlayArea.Enabled = True
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
        picPlayArea.Enabled = False
        Command1.Caption = "New Game"
    End If
End Sub
Sub SetMines()
Dim iPos As Integer

Dim x As Integer
Dim y As Integer
Dim iMines As Integer
Dim iNumCells As Integer
Dim bNotThere As Boolean
    iNumCells = picCell.Count
    iMines = CInt(txtMines)
    If iMines >= picCell.Count Then
        MsgBox "You've specified more mines than squares on the grid! Count reset to 10.", vbExclamation, App.Title
        iMines = 10
        txtMines = "10"
    End If
    ReDim iBombs(iMines)
    Randomize Timer ^ Format(Now, "ss")
    For x = 0 To iMines - 1
        
        Do
            bNotThere = True
            iPos = Int(Rnd * iNumCells)
            For y = 0 To iMines - 1
                If iBombs(y) = iPos Then
                    bNotThere = False
                    Exit For
                End If
            Next y
        Loop Until bNotThere
        iBombs(x) = iPos
    Next x
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    RevealAllCells
    If Command1.Caption = "Stop!" Then
        Command1_Click
    End If
End Sub
Sub RevealAllCells()
Dim x As Integer
    For x = 0 To (itxtX * itxtY) - 1
        RevealCell x
    Next x
End Sub
Sub RevealAllBombs()
Dim x As Integer
    For x = 0 To (itxtX * itxtY) - 1
        RevealBomb x
    Next x
End Sub
Private Sub Form_Load()
Dim x As Integer
    App.Title = "Minesweeper"
    For x = 1 To 4
        picDisp(x).Left = picDisp(x - 1).Left + picDisp(x - 1).Width '+ 10
    Next x
End Sub

Private Sub picCell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If (picCell(Index).Picture = picBlank.Picture) Or _
        (picCell(Index).Picture = picMarked.Picture) Then
        If Button = 1 Then
            If picCell(Index).Picture <> picMarked.Picture Then
                ExamineCell (Index)
                'picCell(Index).Picture = picFlat.Picture
            End If
        Else
            If picCell(Index).Picture = picMarked.Picture Then
                picCell(Index).Picture = picBlank.Picture
            Else
                picCell(Index).Picture = picMarked.Picture
            End If
                
        End If
    End If
    If CountTheBlanks = iMines Then
        Timer1.Enabled = False
        RevealAllBombs
        DoEvents
        MsgBox "Cool - You found 'em all in " & dTimer - 1 & " seconds", vbExclamation, App.Title
        Command1_Click
    End If
End Sub
Private Sub RevealCellsAroundClick(iIn As Integer)
Dim x As Integer
Dim r As Integer
Dim c As Integer
Dim z As Integer
Dim CellsInRad() As Integer
Dim CellsInRadTemp() As Integer
Dim EmptyCells() As Integer
Dim HM As Integer
Dim CellThere As Boolean

    CellsInRad = GiveMeSurroundingCells(iIn)
    c = 1
    While c > 0
        c = 0
        For x = 0 To UBound(CellsInRad) - 1
            HM = HowManyMines(CellsInRad(x))
            If HM = 0 Then
                picCell(CellsInRad(x)).Picture = picFlat.Picture
                ReDim Preserve EmptyCells(c + 1)
                EmptyCells(c) = CellsInRad(x)
                c = c + 1
            Else
                picCell(CellsInRad(x)).Picture = picNo(HM - 1).Picture
            End If
        Next x
        Erase CellsInRad
        If c = 0 Then
            Exit Sub
        End If
        c = 0
        r = 0
        For x = 0 To UBound(EmptyCells) - 1
            Erase CellsInRadTemp
            CellsInRadTemp = GiveMeSurroundingCells(EmptyCells(x))
            For r = 0 To UBound(CellsInRadTemp) - 1
                If picCell(CellsInRadTemp(r)).Picture = picBlank.Picture Then
                    CellThere = False
                    If c > 0 Then
                        For z = 0 To UBound(CellsInRad) - 1
                            If CellsInRadTemp(r) = CellsInRad(z) Then
                                CellThere = True
                                Exit For
                            End If
                        Next z
                    End If
                    If Not CellThere Then
                        ReDim Preserve CellsInRad(c + 1)
                        CellsInRad(c) = CellsInRadTemp(r)
                        c = c + 1
                    End If
                End If
            Next r
        Next x
        Erase EmptyCells
    Wend
    
End Sub
Private Sub ExamineCell(iIn As Integer)
Dim x As Integer
Dim HM As Integer
Dim RA() As Integer
    DoEvents
    HM = HowManyMines(iIn)
    If IsItAMine(iIn) Then
        picCell(iIn).Picture = picNoBomb.Picture
        RevealAllBombs
        Command1_Click
    ElseIf HM = 0 Then
        picCell(iIn).Picture = picFlat.Picture
        RevealCellsAroundClick (iIn)
    Else
        picCell(iIn).Picture = picNo(HM - 1).Picture
    End If
End Sub
Private Sub RevealCell(iIn As Integer)
Dim HM As Integer
    HM = HowManyMines(iIn)
    If IsItAMine(iIn) Then
        picCell(iIn).Picture = picBomb.Picture
    ElseIf HM = 0 Then
        picCell(iIn).Picture = picFlat.Picture
    Else
        picCell(iIn).Picture = picNo(HM - 1).Picture
    End If
End Sub
Private Sub RevealBomb(iIn As Integer)
Dim HM As Integer
    If IsItAMine(iIn) Then
        If picCell(iIn).Picture <> picNoBomb.Picture Then
            picCell(iIn).Picture = picBomb.Picture
        End If
    End If
End Sub

Private Function GiveMeSurroundingCells(iIn As Integer) As Integer()
Dim iD(8) As Integer
Dim iD1() As Integer
Dim x As Integer
Dim z As Integer
Dim RightCol As Boolean
Dim LeftCol As Boolean

    If iIn Mod itxtX = 0 Then
        LeftCol = True
    End If
    If ((iIn + 1 - itxtX) Mod itxtX = 0) And (iIn > 0) Then
        RightCol = True
    End If
    iD(0) = IIf(LeftCol, -1, iIn - 1 - itxtX)  'top left
    iD(1) = iIn - itxtX                        'above
    iD(2) = IIf(RightCol, -1, iIn + 1 - itxtX) 'above right
    iD(3) = IIf(LeftCol, -1, iIn - 1)          'left
    iD(4) = IIf(RightCol, -1, iIn + 1)         'right
    iD(5) = IIf(LeftCol, -1, iIn - 1 + itxtX)  'below left
    iD(6) = iIn + itxtX                        'below
    iD(7) = IIf(RightCol, -1, iIn + 1 + itxtX) 'below right

    z = 0
    For x = 0 To 7
        If iD(x) >= 0 And iD(x) <= ((itxtX * itxtY) - 1) Then
            ReDim Preserve iD1(z + 1)
            iD1(z) = iD(x)
            z = z + 1
        End If
    Next x
    GiveMeSurroundingCells = iD1
End Function
Private Function HowManyMines(iIn As Integer) As Integer
Dim iDir() As Integer
Dim x As Integer

    iDir = GiveMeSurroundingCells(iIn)
    
    HowManyMines = 0
    For x = 0 To UBound(iDir) - 1
        If iDir(x) >= 0 Then
            If IsItAMine(iDir(x)) Then
                HowManyMines = HowManyMines + 1
            End If
        End If
    Next x
End Function
Function IsItAMine(iIn As Integer) As Boolean
Dim x As Integer
    IsItAMine = False
    For x = 0 To UBound(iBombs) - 1
        If iIn = iBombs(x) Then
            IsItAMine = True
            Exit For
        End If
    Next x
End Function

Private Sub Picture1_Click()
    txtX = "10"
    txtY = "10"
    txtMines = "10"
    Command1_Click
End Sub

Private Sub Picture2_Click()
    txtX = "12"
    txtY = "12"
    txtMines = "18"
    Command1_Click
End Sub

Private Sub Picture3_Click()
    txtX = "13"
    txtY = "13"
    txtMines = "72"
    Command1_Click
End Sub

Private Sub Picture4_Click()
    txtX = "13"
    txtY = "13"
    txtMines = "40"
    Command1_Click
End Sub

Private Sub Timer1_Timer()
    SetDisplay (dTimer)
    dTimer = dTimer + 1
End Sub
Sub SetDisplay(iIn As Integer)
Dim sIn As String
Dim x As Integer
    sIn = Format(iIn, "00000")
    For x = 1 To 5
        picDisp(x - 1).Picture = picDig(CInt(Mid(sIn, x, 1))).Picture
    Next x
End Sub
