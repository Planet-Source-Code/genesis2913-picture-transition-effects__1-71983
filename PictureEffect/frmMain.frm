VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example Transition Effect"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   StartUpPosition =   2  'CenterScreen
   Tag             =   "12"
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Circle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   6840
      TabIndex        =   23
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Cross"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   5160
      TabIndex        =   22
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   8520
      TabIndex        =   21
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Fade In"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   3480
      TabIndex        =   20
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Rectangle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   1800
      TabIndex        =   18
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Inset Bottom Right"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   120
      TabIndex        =   17
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Inset Top Left"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   7800
      TabIndex        =   16
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Maze In"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   7800
      TabIndex        =   14
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "2 Bottom 1 Top"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   7800
      TabIndex        =   12
      Top             =   4440
      Width           =   1575
   End
   Begin VB.PictureBox picProg 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   120
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   421
      TabIndex        =   10
      Top             =   5880
      Width           =   6375
      Begin VB.PictureBox picProgress 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   425
         TabIndex        =   11
         Top             =   0
         Width           =   6375
      End
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "2 Right 1 Left"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Split Vertical"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7800
      TabIndex        =   8
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Split Horizontal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7800
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Vertical 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   7800
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Vertical 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   7800
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Vertical 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   7800
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Horizontal 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   7800
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Horizontal 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7800
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdEffect 
      Caption         =   "Horizontal 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picEffect 
      AutoRedraw      =   -1  'True
      Height          =   5685
      Left            =   120
      ScaleHeight     =   375
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   0
      Top             =   120
      Width           =   7560
      Begin VB.PictureBox picBlend 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   19
         Top             =   5160
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Senyor2009"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7800
      TabIndex        =   15
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6600
      TabIndex        =   13
      Top             =   5880
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coding : Senyor2009@yahoo.com.vn"
'Example paint picture and picture transition effect
Option Explicit
'Blend Type
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
Private Const AC_SRC_OVER = &H0
'Blending API
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
'Ellip API
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
'Brush
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long

'-----------------------------------------------------------------------
Dim i
Dim isBusy As Boolean


Private Sub cmdEffect_Click(Index As Integer)
    On Error Resume Next
    Dim lPicture     As StdPicture
    Dim lPictureDraw As StdPicture
    Dim lWidth       As Single
    Dim lHeight      As Single
    Dim BF As BLENDFUNCTION, lBF As Long
    Dim lSplit
    Dim whRate       As Single
    'Set the parameters
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = 128
        .AlphaFormat = 0
    End With
    'Call DoEvents first to avoid windows crash ;))
    lWidth = picEffect.ScaleWidth 'Width need to draw
    lHeight = picEffect.ScaleHeight 'Height need to draw
    'Get the rate
    whRate = (lWidth / 2) / (lHeight / 2)
    'Disable all button
    isBusy = True
    For i = 0 To cmdEffect.UBound
        cmdEffect(i).Enabled = False
    Next
    'Disabled unload
    cmdExit.Enabled = False
    'Set the picture for transformation
    Set lPicture = LoadPicture(Path & "Tag" & StrReverse(Me.Tag) & ".jpg")
    Set lPictureDraw = LoadPicture(Path & "Tag" & Me.Tag & ".jpg")
    With picEffect
        Select Case Index
            Case 0 'Horizontal 1
                For i = 0 To lWidth
                    DoEvents
                    .Cls
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight
                    .PaintPicture lPictureDraw, i, 0, lWidth - i, lHeight, i
                    picProgress.Width = i / lWidth * picProg.ScaleWidth
                Next
            Case 1
                For i = 0 To lWidth
                    DoEvents
                    .Cls
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight
                    .PaintPicture lPictureDraw, i, 0, lWidth - i, lHeight, 0, 0, lWidth - i
                    picProgress.Width = i / lWidth * picProg.ScaleWidth
                Next
            Case 2 'Horizontal 3
                For i = 0 To lWidth / 2
                    DoEvents
                    .Cls
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight
                    .PaintPicture lPictureDraw, 0, 0, lWidth / 2 - i, lHeight, 0, 0, lWidth / 2 - i
                    .PaintPicture lPictureDraw, lWidth / 2 + i, 0, lWidth / 2 - i, lHeight, lWidth / 2 + i, 0, lWidth / 2 - i
                    picProgress.Width = i / (lWidth / 2) * picProg.ScaleWidth
                Next
            Case 3 'Vertical 1
                For i = 0 To lHeight
                    DoEvents
                    .Cls
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight
                    .PaintPicture lPictureDraw, 0, i, lWidth, lHeight - i, 0, i, lWidth, lHeight - i
                    picProgress.Width = i / lHeight * picProg.ScaleWidth
                Next
            Case 4 'Vertical 2
                For i = 0 To lHeight
                    DoEvents
                    .Cls
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight
                    .PaintPicture lPictureDraw, 0, i, lWidth, lHeight - i, 0, 0, lWidth, lHeight - i
                    picProgress.Width = i / lHeight * picProg.ScaleWidth
                Next
            Case 5 'Vertical 3
                For i = 0 To lHeight / 2
                    DoEvents
                    .Cls
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight
                    .PaintPicture lPictureDraw, 0, 0, lWidth, lHeight / 2 - i, 0, 0, lWidth, lHeight / 2 - i
                    .PaintPicture lPictureDraw, 0, lHeight / 2 + i, lWidth, lHeight / 2 + i, 0, lHeight / 2 + i, lWidth, lHeight / 2 + i
                    picProgress.Width = i / (lHeight / 2) * picProg.ScaleWidth
                Next
            Case 6 'Split Horizontal
                lSplit = lWidth / 6
                For i = 0 To lSplit
                    DoEvents
                    .Cls
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight
                    .PaintPicture lPictureDraw, 0, 0, i, lHeight, 0, 0, i
                    .PaintPicture lPictureDraw, lSplit, 0, i, lHeight, lSplit, 0, i
                    .PaintPicture lPictureDraw, lSplit * 2, 0, i, lHeight, lSplit * 2, 0, i
                    .PaintPicture lPictureDraw, lSplit * 3, 0, i, lHeight, lSplit * 3, 0, i
                    .PaintPicture lPictureDraw, lSplit * 4, 0, i, lHeight, lSplit * 4, 0, i
                    .PaintPicture lPictureDraw, lSplit * 5, 0, i, lHeight, lSplit * 5, 0, i
                    picProgress.Width = i / lSplit * picProg.ScaleWidth
                Next
            Case 7 'Split Vertical
                lSplit = lHeight / 6
                For i = 0 To lSplit
                    DoEvents
                    .Cls
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight
                    .PaintPicture lPictureDraw, 0, 0, lWidth, i, 0, 0, lWidth, i
                    .PaintPicture lPictureDraw, 0, lSplit, lWidth, i, 0, lSplit, lWidth, i
                    .PaintPicture lPictureDraw, 0, lSplit * 2, lWidth, i, 0, lSplit * 2, lWidth, i
                    .PaintPicture lPictureDraw, 0, lSplit * 3, lWidth, i, 0, lSplit * 3, lWidth, i
                    .PaintPicture lPictureDraw, 0, lSplit * 4, lWidth, i, 0, lSplit * 4, lWidth, i
                    .PaintPicture lPictureDraw, 0, lSplit * 5, lWidth, i, 0, lSplit * 5, lWidth, i
                    picProgress.Width = i / lSplit * picProg.ScaleWidth
                Next
            Case 8 '2 Right 1 Left
                For i = 0 To lWidth
                    DoEvents
                    .Cls
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight
                    .PaintPicture lPictureDraw, 0, 0, i, lHeight / 3, 0, 0, i, lHeight / 3
                    .PaintPicture lPictureDraw, lWidth - i, lHeight / 3, i, lHeight / 3, lWidth - i, lHeight / 3, i, lHeight / 3
                    .PaintPicture lPictureDraw, 0, (lHeight / 3) * 2, i, lHeight / 3, 0, (lHeight / 3) * 2, i, lHeight / 3
                    picProgress.Width = i / lWidth * picProg.ScaleWidth
                Next
            Case 9 '2 Bottom 1 Top
                For i = 0 To lHeight
                    DoEvents
                    .Cls
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight
                    .PaintPicture lPictureDraw, 0, 0, lWidth / 3, i, 0, 0, lWidth / 3, i
                    .PaintPicture lPictureDraw, lWidth / 3, lHeight - i, lWidth / 3, i, lWidth / 3, lHeight - i, lWidth / 3, i
                    .PaintPicture lPictureDraw, (lWidth / 3) * 2, 0, lWidth / 3, i, (lWidth / 3) * 2, 0, lWidth / 3, i
                    picProgress.Width = i / lHeight * picProg.ScaleWidth
                Next
            Case 10 'Maze In
                'Draw the top first
                For i = 0 To lWidth
                    DoEvents
                    .Cls
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight
                    .PaintPicture lPictureDraw, 0, 0, i, lHeight / 4, 0, 0, i, lHeight / 4
                    picProgress.Width = (i / lWidth * picProg.ScaleWidth) / 6
                Next
                'Next draw vertical right
                For i = 0 To lHeight
                    DoEvents
                    .PaintPicture lPictureDraw, (lWidth / 4) * 3, 0, lWidth / 4, i, (lWidth / 4) * 3, 0, lWidth / 4, i
                    picProgress.Width = picProg.ScaleWidth / 6 + (i / lHeight * picProg.ScaleWidth) / 6
                Next
                'Draw right to left bottom
                For i = 0 To lWidth
                    DoEvents
                    .PaintPicture lPictureDraw, lWidth - i, (lHeight / 4) * 3, i, lHeight / 4, lWidth - i, (lHeight / 4) * 3, i, lHeight / 4
                    picProgress.Width = (picProg.ScaleWidth / 6) * 2 + (i / lWidth * picProg.ScaleWidth) / 6
                Next i
                'Draw bottom to top left
                For i = 0 To (lHeight / 4) * 3
                    DoEvents
                    .PaintPicture lPictureDraw, 0, lHeight - i, lWidth / 4, i, 0, lHeight - i, lWidth / 4, i
                    picProgress.Width = (picProg.ScaleWidth / 6) * 3 + (i / ((lHeight / 4) * 3) * picProg.ScaleWidth) / 6
                Next
                'Draw left to right middle
                For i = 0 To (lWidth / 4) * 3
                DoEvents
                .PaintPicture lPictureDraw, 0, lHeight / 4, i, lHeight / 4, 0, lHeight / 4, i, lHeight / 4
                picProgress.Width = (picProg.ScaleWidth / 6) * 4 + (i / ((lWidth / 4) * 3) * picProg.ScaleWidth) / 6
                Next
                'Draw right to left middle
                For i = (lWidth / 4) To lWidth
                    DoEvents
                    .PaintPicture lPictureDraw, lWidth - i, (lHeight / 4) * 2, i, lHeight / 4, lWidth - i, (lHeight / 4) * 2, i, lHeight / 4
                    picProgress.Width = (picProg.ScaleWidth / 6) * 5 + (i / lWidth * picProg.ScaleWidth) / 6
                Next
            Case 11 'Inset Top Left
                For i = 0 To IIf(lWidth > lHeight, lHeight, lWidth)
                    DoEvents
                    .Cls
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight
                    .PaintPicture lPictureDraw, 0, 0, (lWidth / lHeight) * i, i, 0, 0, (lWidth / lHeight) * i, i
                    picProgress.Width = (i / IIf(lWidth > lHeight, lHeight, lWidth)) * picProg.ScaleWidth
                Next
            Case 12 'Inset Bottom Right
                .Cls
                .PaintPicture lPicture, 0, 0, lWidth, lHeight
                For i = 0 To IIf(lWidth > lHeight, lHeight, lWidth)
                    DoEvents
                    .PaintPicture lPictureDraw, lWidth - ((lWidth / lHeight) * i), lHeight - i, (lWidth / lHeight) * i, i, lWidth - ((lWidth / lHeight) * i), lHeight - i, (lWidth / lHeight) * i, i
                    picProgress.Width = (i / IIf(lWidth > lHeight, lHeight, lWidth)) * picProg.ScaleWidth
                Next
            Case 13 'Rectangle
                For i = 0 To lHeight / 2
                    DoEvents
                    .Cls
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight
                    .PaintPicture lPictureDraw, lWidth / 2 - whRate * i, lHeight / 2 - i, (whRate * i) * 2, i * 2, lWidth / 2 - whRate * i, lHeight / 2 - i, (whRate * i) * 2, i * 2
                    picProgress.Width = (i / (lHeight / 2)) * picProg.ScaleWidth
                Next
            Case 14 'Fade in
                picBlend.Move 0, 0, .ScaleWidth, .ScaleHeight
                For i = 0 To 255
                    DoEvents
                    .Cls 'Clean up
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight 'Draw first picture
                    picBlend.PaintPicture lPictureDraw, 0, 0, lWidth, lHeight 'Draw second picture
                    BF.SourceConstantAlpha = i 'Set the alpha
                    RtlMoveMemory lBF, BF, 4 'Convert to long
                    'Now blending two picture ;))
                    AlphaBlend .hdc, 0, 0, lWidth, lHeight, picBlend.hdc, 0, 0, lWidth, lHeight, lBF
                    picProgress.Width = (i / 255) * picProg.ScaleWidth
                Next
            Case 15 'Cross
                For i = Int(lHeight / 2) To 0 Step -1
                    DoEvents
                    .Cls
                    .PaintPicture lPictureDraw, 0, 0, lWidth, lHeight
                    'Draw top left
                    .PaintPicture lPicture, 0, 0, whRate * i, i, 0, 0, whRate * i, i
                    'Draw bottom left
                    .PaintPicture lPicture, 0, lHeight - i, whRate * i, i, 0, lHeight - i, whRate * i, i
                    'Draw top right
                    .PaintPicture lPicture, lWidth - whRate * i, 0, whRate * i, i, lWidth - whRate * i, 0, whRate * i, i
                    'Draw bottom right
                    .PaintPicture lPicture, lWidth - whRate * i, lHeight - i, whRate * i, i, lWidth - whRate * i, lHeight - i, whRate * i, i
                Next
            Case 16 'Circle
                Dim hCircle As Long
                Dim mBrush  As Long
                Dim dCheo   As Long
                'Tinh duong cheo cua buc anh
                dCheo = Sqr(lWidth * lWidth + lHeight * lHeight) 'Pytago ;))
                dCheo = dCheo / 2 'Get 1/2
                'Get picture
                mBrush = CreatePatternBrush(lPictureDraw.Handle)
                For i = 0 To dCheo
                    'Draw circle
                    hCircle = CreateEllipticRgn(0, 0, i * 2, i * 2)
                    'Move it
                    OffsetRgn hCircle, lWidth / 2 - i, lHeight / 2 - i
                    DoEvents
                    .Cls
                    .PaintPicture lPicture, 0, 0, lWidth, lHeight
                    'Paint it
                    hCircle = FillRgn(.hdc, hCircle, mBrush)
                    'Clean up
                    DeleteObject hCircle
                    picProgress.Width = i / dCheo * picProg.ScaleWidth
                Next
                'Remove mBrush
                DeleteObject mBrush
        End Select
    End With
    
    'Get new picture
    Me.Tag = StrReverse(Me.Tag)
    Me.Refresh
    'Redraw the picture
    If Index = 6 Or Index = 7 Or Index = 10 Or Index = 11 Or Index = 12 Or Index = 13 Then
        picEffect.Cls
        picEffect.PaintPicture lPictureDraw, 0, 0, lWidth, lHeight
    End If
    'Refresh the form
    picEffect.Refresh
    'Enabled all button
    For i = 0 To cmdEffect.UBound
        cmdEffect(i).Enabled = True
    Next
    'Set the label percent
    'For cheating ;))
    picProgress.Width = picProg.ScaleWidth
    lblPercent.Caption = "100%"
    isBusy = False
    'Now you can exit ;)
    cmdExit.Enabled = True
End Sub

'Get the application path
Private Function Path() As String
    Path = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
End Function

'Unload me
Private Sub cmdExit_Click()
    Unload Me
End Sub

'You can't open multiple
Private Sub Form_Initialize()
    If App.PrevInstance Then End
End Sub

Private Sub Form_Load()
    MsgBox "Coding : Senyor@yahoo.com.vn" & vbNewLine & "Example paint picture and picture transition effect !", vbInformation + vbOKOnly, "Readme"
    lblPercent.Caption = "100%"
    picEffect.PaintPicture LoadPicture(Path & "Tag" & Me.Tag & ".jpg"), 0, 0, picEffect.ScaleWidth, picEffect.ScaleHeight
End Sub

'You can exit when anything done ;)
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If isBusy = True Then Cancel = 1
End Sub

Private Sub picProgress_Resize()
    On Error Resume Next
    'Get the percent for label
    lblPercent.Caption = Format(picProgress.Width / picProg.ScaleWidth, "00" & "%")
End Sub
