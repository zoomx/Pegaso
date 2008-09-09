Attribute VB_Name = "FileFunctions"
Option Explicit

Public Sub ParseLine(Linea As String, Vettore() As String, nElementi As Long, RecordDelimiter As String)
    'Prende in ingresso una linea di dati e restituisce un vettore con i dati separati
    Vettore = Split(Linea, RecordDelimiter)
    nElementi = UBound(Vettore)
End Sub

Public Sub ParseData(stringa As String)
    Dim i As Integer
    Static Buffer As String
    Dim Linea As String
    
'    Dim j As Integer
'    j = FreeFile
'    Open App.Path & "\" & "file.txt" For Append As #j
'    Print #j, stringa
'    Close j
    
    Buffer = Buffer + stringa
    i = InStr(Buffer, vbLf)
    'Debug.Print "i="; i
    If i <> 0 Then
        Linea = Left$(Buffer, i)
        'parse linea
        Debug.Print Linea
        Buffer = Mid(Buffer, i + 1)
   End If
    

End Sub
