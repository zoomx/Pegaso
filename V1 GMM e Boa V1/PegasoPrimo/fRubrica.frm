VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form fRubrica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Poseidon - Rubrica"
   ClientHeight    =   3900
   ClientLeft      =   4470
   ClientTop       =   4515
   ClientWidth     =   6435
   Icon            =   "fRubrica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   "Rubrica"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Rubrica"
      Top             =   120
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "&Annulla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "fRubrica.frx":0442
      Height          =   3600
      Left            =   120
      OleObjectBlob   =   "fRubrica.frx":0456
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "fRubrica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    LeggiRubrica
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    fModem.Show
End Sub

Private Sub CmdAnnulla_Click()
    Unload Me
    fModem.Show
End Sub

Private Sub DBGrid1_DblClick()
    'If DBGrid1.Col = 0 Then
        DBGrid1.Col = 0
        fModem.txtNumero.Text = DBGrid1.Text
    'End If
    CmdAnnulla_Click
End Sub

Public Sub LeggiRubrica()
    'Carica la rubrica
    Dim Path As String
    Dim NomeFile As String
    Dim nfile As Integer
    Dim ws1 As Workspace
    Dim db1 As Database
    Dim td1 As TableDef
    Dim fl1, fl2 As Field
    Dim rs1 As Recordset
    
    Path = sGetAppPath()
    nfile = FreeFile

    'Cerca la rubrica
    NomeFile = Path + "Rubrica.mdb"
    Err.Clear
    On Local Error Resume Next
    Open NomeFile For Input As #nfile
    If Err.Number = 53 Then
        'La rubrica non c'è, ne crea una vuota
        NomeFile = "ERRORE!" + vbCrLf
        NomeFile = NomeFile + "Impossibile trovare la rubrica."
        MsgBox (NomeFile)
        'ScriviErrore (NomeFile)
        'Exit Sub
        
        ' Create Work space
        ' create Data base
        ' create Table
        
        Set ws1 = DBEngine.Workspaces(0)
        Set db1 = ws1.CreateDatabase(App.Path & "\Rubrica.mdb", dbLangGeneral, dbVersion30)
        Set td1 = db1.CreateTableDef("Rubrica")
        
        'Create Fields in table 1 (Only 2)

        Set fl1 = td1.CreateField("Telefono", dbText)
        td1.Fields.Append fl1
        Set fl2 = td1.CreateField("Descrizione", dbText)
        td1.Fields.Append fl2

        ' Append fields to Data base
        
        db1.TableDefs.Append td1
        
        ' Open RecordSet

        Set rs1 = db1.OpenRecordset("Rubrica", dbOpenTable)

        'Add Data to each Field in record set
        rs1.AddNew
        rs1.Fields("Telefono") = "095123456"
        rs1.Fields("DEscrizione") = "Fake"

        ' Update or it will be lost!

        rs1.Update
        'Chiude il database
        db1.Close
    ElseIf Err.Number = 0 Then
    
    Else
        ErrHandler
        Close nfile
        Exit Sub
    End If
    Close nfile
    'NomeFile = Path + "Rubrica.txt"
    'NomeFile = Path + "Rubrica.csv"
    'NomeFile = Path + "Rubrica.xls"
    NomeFile = Path + "Rubrica.mdb"

    'legge la rubrica
    'Data1.DefaultType = 2
    'Data1.DefaultType = 1
    'Data1.DatabaseName = Path
    
    'Data1.Connect = "Text;" '& Path
    'Data1.Connect = "dBASE III;"
    'Data1.Connect = "Excel 5.0;" & NomeFile & ";"
    Data1.Connect = "Access" '& NomeFile
    Data1.RecordsetType = 1
    Data1.DatabaseName = NomeFile
    
    'Data1.DatabaseName = "Rubrica"
    Data1.RecordSource = "SELECT * FROM Rubrica"
    
    Data1.Refresh
    
End Sub
