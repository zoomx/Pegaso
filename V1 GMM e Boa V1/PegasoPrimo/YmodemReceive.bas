Attribute VB_Name = "YmodemReceive"
Option Explicit

Public Const STARTY As String = "C"
Public Const SOH As Byte = 1
Public Const STX As Byte = 2
Public Const EOT As Byte = 4
Public Const ACK As Byte = 6
Public Const NAK As Byte = 15
Public Const CAN As Byte = 18
Public Const MAXTRY As Byte = 3
Public Const xyBufferSize As Integer = 1024
Public Const Port As Byte = 1
Public Const LIMIT As Integer = 20
Dim NCGbyte As Byte
  
  
Public Function YmodemReceiveFile(Path As String) As String

  Dim Buffer(1024) As Byte
  Dim TheByte      As Byte
  Dim BufferSize   As Integer
  Dim ErrorFlag    As Boolean
  Dim EOTflag      As Boolean
  Dim FirstPacket  As Integer
  Dim Code         As Integer
  Dim FileNbr      As Integer
  Dim Packet       As Integer
  Dim PacketNbr    As Integer
  Dim i            As Integer
  Dim Flag         As Integer
  Dim FileBytes    As Long
  Dim AnyKey       As String
  Dim Message      As String
  Dim Temp         As String
  Dim RXYmodem As Boolean
  Dim BatchFlag As Boolean
  
  BatchFlag = True
  ErrorFlag = False
  EOTflag = False
  NCGbyte = Asc(STARTY) 'C
  'CALL WriteMsg("XYMODEM Receive: Waiting for Sender ")
  
  '  'clear comm port
  'Code = SioRxFlush(Port)

  'Send NAKs or 'C's
  If Not RxStartup(Port, NCGbyte) Then
    RXYmodem = False
    Exit Function
  End If

  'open file unless BatchFlag is on
  If BatchFlag Then
    FirstPacket = 0
  Else
    FirstPacket = 1
    'Open file for write
    FileNbr = FreeFile
    Open Filename For Binary Access Write As FileNbr
    Debug.Print "Opening "; Filename
  End If



  'get each packet in turn
  For Packet = FirstPacket To 32767
    'user aborts ?
'    IF AnyKey$ = STR$(%CAN) THEN
'      Call TxCAN(Port)
'      Call WriteMsg("*** Canceled by USER ***")
'      RxyModem = False
'      Exit Function
'    End If

    'issue message
    Message = "Packet " + Str$(Packet)
'    Call WriteMsg(Message)
    Debug.Print Message
    PacketNbr = Packet And 255
    'get next packet
    If Not RxPacket(Port, Packet, Buffer(), BufferSize, NCGbyte, EOTflag) Then
      RXYmodem = False
      YmodemReceiveFile = "Error"
      Exit Function
    End If
    'packet 0 ?
    If Packet = 0 Then
      'name & date packet
      If Buffer(0) = 0 Then
'        Call WriteMsg("Batch transfer complete")
        Debug.Print "Batch transfer complete"
        RXYmodem = True
        YmodemReceiveFile = "OK"
        Exit Function
      End If
      'construct filename
      i = 0
      Filename = ""
      Do
        TheByte = Buffer(i)
        If TheByte = 0 Then
          Exit Do
        End If
        Filename = Filename + Chr$(TheByte)
        i = i + 1
      Loop
      Debug.Print Filename
      'get file size
      i = i + 1
      Temp$ = ""
      Do
        TheByte = Buffer(i)
        If TheByte = 0 Then
          Exit Do
        End If
        Temp$ = Temp$ + Chr$(TheByte)
        i = i + 1
      Loop
      FileBytes = Val(Temp$)
      Debug.Print FileBytes
    End If
    'all done if EOT was received
    If EOTflag Then
      Close FileNbr
'      Call WriteMsg("Transfer completed")
      Debug.Print "Transfer completed"
      RXYmodem = True
      YmodemReceiveFile = "OK"
      Exit Function
    End If
    'process the packet
    If Packet = 0 Then
      'open file using filename in packet 0
      FileNbr = FreeFile
      Open Filename For Binary Access Write As FileNbr
'      Print "Opening "; filename
      Debug.Print "Opening "; Filename
      'must restart after packet 0
      Flag = RxStartup(Port, NCGbyte)
    Else
      'Packet > 0  ==> write Buffer
      For i = 0 To BufferSize - 1
        Put FileNbr, , Buffer(i)
      Next i
    End If
  Next Packet

RxyM_EXIT:
  Close FileNbr
  Exit Function

RxyTrap:
  Select Case Err
    Case 53
      Message = "Cannot open " + Filename + " for write"
      'Call WriteMsg(Message)
      Debug.Print Message
    Case Else
      'Print "RX Error: ("; Err; ")"
      Debug.Print "RX Error: ("; Err; ")"
    End Select

    RXYmodem = False
    Resume RxyM_EXIT


End Function

Public Function RxPacket(ByVal Port As Integer, _
                  ByVal PacketNbr As Integer, _
                        Buffer() As Byte, _
                        PacketSize As Integer, _
                  ByVal NCGbyte As Byte, _
                        EOTflag As Boolean)

  'Port      : Port # [0..3)
  'PacketNbr : Packet # [0,1,2,...)
  'PacketSize: Packet size [128,1024) {returned}
  'NCGbyte   : NAK, "C", or "G"
  'EOTflag   : EOT was received       {returned}

  Dim i            As Integer
  Dim CheckSum     As Long
  Dim RxCheckSum   As Long
  Dim RxCheckSum1  As Long
  Dim RxCheckSum2  As Long
  Dim Attempt      As Integer
  Dim Code         As Integer
  Dim PacketType   As Integer
  Dim RxPacketNbr  As Integer
  Dim RxPacketNbrC As Integer

  PacketNbr = PacketNbr And 255

  For Attempt = 1 To MAXTRY
    'wait FOR SOH / STX
    Code = SioGetc(Port, 2)
    
    If Code = -1 Then
      'Print "Timed out waiting for sender"
      Debug.Print "Timed out waiting for sender"
      RxPacket = False
      Exit Function
    End If
    Select Case Code
      Case SOH
        '128 byte buffer incoming
        PacketType = SOH
        PacketSize = 128
      Case STX
        '1024 byte buffer incoming
        PacketType = STX
        PacketSize = 1024
      Case EOT
        'all packets have been sent
        Code = SioPutc(Port, ACK)
        EOTflag = True
        RxPacket = True
        Exit Function
      Case CAN
        'sender has canceled !
        Debug.Print "Canceled by remote"
        RxPacket = False
      Case Else
        'error !
        Debug.Print "Expecting SOH/STX/EOT/CAN not "; Code
        RxPacket = False
    End Select

    'receive packet #
    Code = SioGetc(Port, 1)
    If Code = -1 Then
      Debug.Print "Timed out waiting for packet #"
      Exit Function
    End If
    RxPacketNbr = Code And 255

    'receive 1's complement
    Code = SioGetc(Port, 1)
    If Code = -1 Then
      Debug.Print "Timed out waiting for complement of packet #"
      RxPacket = False
      Exit Function
    End If
    RxPacketNbrC = Code And 255

    'receive data
    CheckSum = 0
    For i = 0 To PacketSize - 1
      Code = SioGetc(Port, 1)
      If Code = -1 Then
        Debug.Print "Timed out waiting for data for packet #"
        RxPacket = False
        Exit Function
      End If
      Buffer(i) = Code
      'compute CRC or checksum
      If NCGbyte <> NAK Then
        CheckSum = UpdateCRC(Code, CheckSum)
      Else
        CheckSum = (CheckSum + Code) And 255
      End If
    Next i

    'receive CRC/checksum
    If NCGbyte <> NAK Then
      'receive 2 byte CRC
      Code = SioGetc(Port, 1)
      If Code = -1 Then
        Debug.Print "Timed out waiting for 1st CRC byte"
        Exit Function
      End If
      RxCheckSum1 = Code And 255
      Code = SioGetc(Port, 1)
      If Code = -1 Then
        Debug.Print "Timed out waiting for 2nd CRC byte"
        RxPacket = False
        Exit Function
      End If
      RxCheckSum2 = Code And 255
      RxCheckSum = (256 * RxCheckSum1) Or RxCheckSum2
    Else
      'receive one byte checksum
      Code = SioGetc(Port, 1)
      If Code = -1 Then
        Debug.Print "Timed out waiting for checksum"
        RxPacket = False
        Exit Function
      End If
      RxCheckSum = Code And 255
    End If

    'don't send ACK IF "G"
    If NCGbyte = Asc("G") Then
      RxPacket = True
      Exit Function
    End If

    'packet # and checksum OK ?
    If (RxCheckSum = CheckSum) And (RxPacketNbr = PacketNbr) Then
      'ACK the packet
      Code = SioPutc(Port, ACK)
      RxPacket = True
      Exit Function
    End If

    'bad packet
    If RxCheckSum = CheckSum Then
      Debug.Print "Bad Packet. Received "; RxPacketNbr; ", expected "; PacketNbr
    Else
      Debug.Print "Bad Checksum. Received "; Hex$(RxCheckSum); _
            ", expected "; Hex$(CheckSum)
    End If
    Code = SioPutc(Port, NAK)
  Next Attempt

  'can't receive packet
  Debug.Print "RX packet timeout"
  RxPacket = False

End Function


Public Function YmodemRx(ByVal Port As Integer, _
                        Filename As String, _
                  ByVal NCGbyte As Byte)

  Dim AnyKey As String

  YmodemRx = True
  Do
    AnyKey$ = INKEY$
    If AnyKey$ <> "" Then
      Call WriteMsg("Aborted by user")
      Exit Do
    End If
    Call WriteMsg("Ready for next file")
    Filename = ""
    If Not RXYmodem(Port, Filename, NCGbyte, True) Then
      YmodemRx = False
      Exit Function
    End If
    'empty filename ?
    If Filename = "" Then
      Exit Function
    End If
  Loop

End Function

Public Function RxStartup(ByVal Port As Integer, _
                   ByVal NCGbyte As Byte)
  Dim i       As Integer
  Dim Code    As Integer
  Dim Code2   As Integer
  Dim TheByte As Byte
  Dim AnyKey  As String

  'clear Rx buffer
  'Code = SioRxFlush(Port)

  'Send NAKs or "C"s
  For i = 1 To LIMIT
'    AnyKey$ = INKEY$
'    If AnyKey$ <> "" Then
'      Debug.Print "Canceled by user"
'      RxStartup = False
'      Exit Function
'    End If
    'stop attempting CRC after 1st 4 tries
    If (NCGbyte <> NAK) And (i = 5) Then NCGbyte = NAK
    'tell sender that I am ready to receive
    Code = SioPutc(Port, NCGbyte)
'    Code = SioGetc(Port, 2)
'    If Code <> -1 Then
    If fMain.MSComm1.InBufferCount <> 0 Then
'      'no error -- must be incoming byte -- push byte back onto queue !
'      Code2 = SioUnGetc(Port, Code)
      RxStartup = True
      Exit Function
    End If
  Next i

  'no response
  Debug.Print "No response from sender"
  RxStartup = False

End Function


Public Function SioGetc(Port As Integer, TimeOut As Byte) As Integer
    SioGetc = InputComTimeOutBin3(TimeOut)
End Function

Public Function SioPutc(Port As Integer, Char As Byte) As Integer
    fMain.MSComm1.Output = Chr(Char)
    SioPutc = 1
End Function
