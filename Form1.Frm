VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "ProgressBar using only a PictureBox"
   ClientHeight    =   1785
   ClientLeft      =   2640
   ClientTop       =   3225
   ClientWidth     =   6690
   FillColor       =   &H000040C0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1785
   ScaleWidth      =   6690
   Begin MSComDlg.CommonDialog dlg 
      Left            =   1920
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Color"
   End
   Begin VB.Frame Frame4 
      Caption         =   "% Align"
      Height          =   615
      Left            =   3720
      TabIndex        =   11
      Top             =   240
      Width           =   2775
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":000C
         Left            =   120
         List            =   "Form1.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "ForeColor"
      Height          =   615
      Left            =   2520
      TabIndex        =   9
      Top             =   240
      Width           =   1095
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "BackColor"
      Height          =   615
      Left            =   1320
      TabIndex        =   7
      Top             =   240
      Width           =   1095
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "LineColor"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1095
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "5000"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   6435
      TabIndex        =   2
      Top             =   960
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "ProgressBar Deluxe v1.1"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Code example by Rudy Alex Kohn
' Use as you like, but plz credit me for it =)
' rudyalexkohn@hotmail.com
Option Explicit

Private Sub Command1_Click()
' Additional notes :
' - The picturebox should have autoredraw enabled, then you can remove the AutoRedraw check
' If using a long. try val()

  Command1.Enabled = False
  Dim iValue As Integer
  Dim i As Integer                                    ' Loop counter

  If IsNumeric(Text1) Then
    iValue = Text1
  Else
    iValue = 5000
  End If
  Select Case iValue
  Case Is <= 0, Is >= 30000
    iValue = 5000
  End Select

  Dim nBackColor As Long
  Dim nLineColor As Long
  Dim nForeColor As Long
  Dim nAlign As Long

  nLineColor = Label1.BackColor
  nBackColor = Label2.BackColor
  nForeColor = Label3.BackColor
  Select Case Combo1.Text
  Case "Center"
    nAlign = vbCenter
  Case "Right"
    nAlign = vbRightJustify
  Case Else
    nAlign = vbLeftJustify
  End Select
  If Not (Picture1.AutoRedraw) Then Picture1.AutoRedraw = True  ' Else it flickers like hell.
  For i = 1 To iValue Step 10
    ' This is where the progress bar is updated.
    ' Add you're code being executed here, or modify it to suit you'r needs
    DrawPercent Picture1, i / iValue * 100, nForeColor, nLineColor, nBackColor, nAlign, lblTitle ' Progress bar update
    DoEvents                                        ' Don't remove
  Next
  Picture1.Cls                                        ' Clear it after use
  Picture1.BackColor = &H8000000F                     ' Sets backcolor to 'Button Face' (default)
  Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
  Combo1.Text = "Center"
End Sub

Private Sub Label1_Click()
  With dlg
    .Color = Label1.BackColor
    .ShowColor
    Label1.BackColor = .Color
  End With
End Sub

Private Sub Label2_Click()
  With dlg
    .Color = Label2.BackColor
    .ShowColor
    Label2.BackColor = .Color
  End With
End Sub

Private Sub Label3_Click()
  With dlg
    .Color = Label3.BackColor
    .ShowColor
    Label3.BackColor = .Color
  End With
End Sub
