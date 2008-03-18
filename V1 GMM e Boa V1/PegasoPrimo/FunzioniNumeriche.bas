Attribute VB_Name = "FunzioniNumeriche"
Option Explicit

Public Function Val2(Valore As Variant) As Variant
'Simile alla val ma per separatore decimale usa sia il
'punto che la virgola
    Dim ip As Integer
    Dim iv As Integer
    Dim lStringa As Integer
    Dim Temp As Variant
    Dim Stringa As String
    
    Stringa = CStr(Valore)
    'C'è il punto?
    ip = InStr(Stringa, ".")
    'C'è la virgola?
    iv = InStr(Stringa, ",")
    lStringa = Len(Stringa)
    If iv <> 0 Then 'Se c'è la virgola la sostituisce col punto
        Stringa = Left(Stringa, iv - 1) + "." + Right(Stringa, lStringa - iv)
        ip = iv
    End If
    Temp = Val(Stringa)
    'If ip <> 0 And iv <> 0 Then
    'Se ci sono tutte e due?
    Val2 = Temp
End Function


Public Function bytes2long(Stringa As String) As Long
'converte una stringa rappresentante un numero long in binario
'(littel endian, basso-alto) nel numero stesso
    Dim lStringa As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Lungo As Long
    Dim a As String
    On Error GoTo GestErr
    'StampaAscii (Stringa)
    lStringa = Len(Stringa)
    If lStringa > 4 Then
        Stringa = Left(Stringa, 4)
        lStringa = Len(Stringa)
    End If
    Lungo = 0
    'For i = lstringa To 1 Step -1
    For i = 1 To lStringa
        a = Mid(Stringa, i, 1)
        j = Asc(a)
        Lungo = Lungo + j * 256 ^ (i - 1)
        
    Next
    bytes2long = Lungo
    Exit Function
GestErr:
    If Err.Number = 6 Then
        bytes2long = 2147483647
    End If
    
End Function

Public Function String2long(Stringa As String) As Long
    Dim lStringa As Integer
    Dim Lungo As Long
    
'    lStringa = Len(Stringa)
'    If lStringa <> 4 Then
'        Messaggio = "La lunghezza del numero è errata ->" + Str(lStringa) + " invece di 4"
'        MsgBox (Messaggio)
'    End If

    Stringa = SwapString(Stringa)
    Lungo = bytes2long(Stringa)
    String2long = Lungo
End Function

Public Function SwapString(Stringa As String) As String
    Dim lStringa As Long
    Dim Dummy As String
    Dim i As Long
    lStringa = Len(Stringa)
    'Capovolge la stringa
    Dummy = ""
    For i = lStringa To 1 Step -1
        Dummy = Dummy + Mid(Stringa, i, 1)
    Next
    SwapString = Dummy
End Function
