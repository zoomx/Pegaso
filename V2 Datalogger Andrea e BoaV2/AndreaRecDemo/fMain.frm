VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Pegaso On Line"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10620
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   480
      TabIndex        =   17
      Top             =   4440
      Width           =   3135
      Begin VB.PictureBox pDepth 
         Height          =   615
         Left            =   1560
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Depth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Current meter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   5520
      TabIndex        =   10
      Top             =   240
      Width           =   4695
      Begin VB.PictureBox pFlow 
         Height          =   615
         Left            =   1560
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.PictureBox pPolar 
         Height          =   3910
         Left            =   240
         Picture         =   "fMain.frx":0CCA
         ScaleHeight     =   3855
         ScaleWidth      =   4065
         TabIndex        =   11
         Top             =   1080
         Width           =   4120
      End
   End
   Begin VB.CommandButton bTestLine 
      Caption         =   "Test Line"
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton bConnect 
      Caption         =   "&Connect"
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Water"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.PictureBox pTurbidity 
         Height          =   615
         Left            =   2280
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   16
         Top             =   3120
         Width           =   1335
      End
      Begin VB.PictureBox pOxygen 
         Height          =   615
         Left            =   2280
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   15
         Top             =   2400
         Width           =   1335
      End
      Begin VB.PictureBox ppH 
         Height          =   615
         Left            =   2280
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.PictureBox pCond 
         Height          =   615
         Left            =   2280
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.PictureBox pTemperature 
         Height          =   615
         Left            =   2280
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Turbidity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Oxygen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "pH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Conducibility"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Temperature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton bEnd 
      Caption         =   "&End"
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   5760
      Width           =   1215
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private LCDtemp As New mcLCD
Private LCDcond As New mcLCD
Private LCDpH As New mcLCD
Private LCDflow As New mcLCD
Private LCDoxy As New mcLCD
Private LCDturb As New mcLCD
Private LCDdepth As New mcLCD

Private Sub bEnd_Click()
    End
End Sub

Private Sub bTestLine_Click()
    DrawLine 250, 0
    DrawLine 500, 90
    DrawLine 1000, 180
    DrawLine 1500, 270
End Sub

Private Sub Form_Load()
    With LCDtemp
        .NewLCD pTemperature
        .BackColor = vbBlack
        .ForeColor = vbRed
        .Caption = "20.5"
    End With
    With LCDtemp
        .NewLCD pCond
        .BackColor = vbBlack
        .ForeColor = vbRed
        .Caption = "55"
    End With
    With LCDtemp
        .NewLCD ppH
        .BackColor = vbBlack
        .ForeColor = vbRed
        .Caption = "8.0"
    End With
    With LCDflow
        .NewLCD pFlow
        .BackColor = vbBlack
        .ForeColor = vbRed
        .Caption = "20.5"
    End With
    With LCDoxy
        .NewLCD pOxygen
        .BackColor = vbBlack
        .ForeColor = vbRed
        .Caption = "3.6"
    End With
    With LCDturb
        .NewLCD pTurbidity
        .BackColor = vbBlack
        .ForeColor = vbRed
        .Caption = "15.3"
    End With
    With LCDdepth
        .NewLCD pDepth
        .BackColor = vbBlack
        .ForeColor = vbRed
        .Caption = "22.5"
    End With

End Sub

Sub DrawLine(ByVal length As Single, ByVal Angle As Single)
    Dim Pi As Double
    Dim cx As Single
    Dim cy As Single
    Dim xp As Single
    Dim yp As Single
    Dim Rx As Single
    Dim Ry As Single
    Dim rxg As Single
    Dim ryg As Single
    
    
    Pi = 4 * Atn(1)
    cx = pPolar.Width / 2
    cy = pPolar.Height / 2
    
    'Angle is in Degrees
    Angle = Angle + 180
    Angle = Angle Mod 360
    
    Angle = Angle * Pi / 180
    xp = 0
    yp = Abs(length)
    Rx = xp * Cos(Angle) - yp * Sin(Angle)
    Ry = xp * Sin(Angle) + yp * Cos(Angle)
    rxg = cx + Rx
    ryg = cy - Ry
    
    ryg = pPolar.Height - ryg
    
    pPolar.ForeColor = vbRed
    pPolar.DrawWidth = 5
    pPolar.Line (cx, cy)-(rxg, ryg)
    
    'Line (cx, cy)-(rxg, ryg)
    
'    ' if negative length go back to start position
'    If length < 0 Then
'        Me.CurrentX = cx
'        Me.CurrentY = cy
'    End If
'

End Sub

