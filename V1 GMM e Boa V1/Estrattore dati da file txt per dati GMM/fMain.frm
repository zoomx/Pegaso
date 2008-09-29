VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Estrattore"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton bFileOpen 
      Caption         =   "File into search"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox TextToSearch 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Text            =   "CTD DATA EMPTY!"
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Lline 
      Caption         =   "Line"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Text to search"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CD As New clsDialog
Dim FileIn As String
Dim Fileout As String

Private Sub bFileOpen_Click()
FileIn = CD.ShowOpen(Me.hWnd, "", CD.FileName, "Text (*.txt)|*.txt")

End Sub

Private Sub bStart_Click()
    Dim Stringa As String
    Open FileIn For Input As #1
    Open Fileout For Output As #2
    Do
        Line Input #1, Stringa
        Lline.Caption = Stringa
        
        If InStr(Stringa, TextToSearch.Text) <> 0 Then
            Print #2, Stringa
        End If
    Loop Until EOF(1)
    Close 1
    Close 2
    Lline.Caption = "DONE!"
End Sub

Private Sub Form_Load()
    Fileout = App.Path + "\estratto.txt"
End Sub
