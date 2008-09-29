Attribute VB_Name = "Globals"
Public Type OPENFILENAME
  lStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  lpstrFilter       As String
  lpstrCustomFilter As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  lpstrFile         As String
  nMaxFile          As Long
  lpstrFileTitle    As String
  nMaxFileTitle     As Long
  lpstrInitialDir   As String
  lpstrTitle        As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  lpstrDefExt       As String
  lCustData         As Long
  lpfnHook          As Long
  lpTemplateName    As String
End Type

Public Const OFN_ALLOWMULTISELECT   As Long = &H200
Public Const OFN_CREATEPROMPT       As Long = &H2000
Public Const OFN_EXPLORER           As Long = &H80000
Public Const OFN_EXTENSIONDIFFERENT As Long = &H400
Public Const OFN_FILEMUSTEXIST      As Long = &H1000
Public Const OFN_HIDEREADONLY       As Long = &H4
Public Const OFN_LONGNAMES          As Long = &H200000
Public Const OFN_NOCHANGEDIR        As Long = &H8
Public Const OFN_NODEREFERENCELINKS As Long = &H100000
Public Const OFN_OVERWRITEPROMPT    As Long = &H2
Public Const OFN_PATHMUSTEXIST      As Long = &H800
Public Const OFN_READONLY           As Long = &H1

