VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Caller"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   15180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bEnd 
      Caption         =   "&End"
      Height          =   495
      Left            =   13800
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton bStart 
      Caption         =   "&Start"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lMonitor 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   14655
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bEnd_Click()
    End
End Sub

Private Sub bStart_Click()
    Dim FileName As String
    Dim LineBuffer As String
    Dim Linea As String
    Dim LineCount As Long
    Dim LineLenght As Long
    Dim Lines As Integer
    
    FileName = App.Path + "\" + "1.log"
    Open FileName For Input As #1
    Line Input #1, LineBuffer
    Close 1
    FileName = App.Path + "\" + "1.txt"
    Open FileName For Output As #1
    LineLenght = Len(LineBuffer)
    Lines = LineLenght / 60
    Debug.Print LineLenght; "->"; Lines; " lines"
    For LineCount = 0 To Lines - 1
        Linea = Mid(LineBuffer, LineCount * 60 + 1, 60)
        'Print #1, Linea
        Linea = ParseLine(Linea)
        Print #1, Linea
    Next LineCount
    Close 1
End Sub
