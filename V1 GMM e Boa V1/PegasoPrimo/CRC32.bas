Attribute VB_Name = "CRC32"
Option Explicit
Private malngLookup(0 To 255) As Long
Private Const CRC_BUFFER_LENGTH = 1024

Public Function GetCrc32(ByVal strFile As String) As Long
    Dim lngCrc32 As Long
    Dim i As Long
    Dim lngFilePtr As Long
    Dim lngFileLength As Long
    Dim abytBuffer(CRC_BUFFER_LENGTH) As Byte
    Dim lngBufferPtr As Long
    Dim lngBufferLength As Long
    
    CreateCrc32Lookup
    lngCrc32 = &HFFFFFFFF
    
    For i = 1 To Len(strFile)
        lngCrc32 = (Int(lngCrc32 / 256) And &HFFFFFF) Xor (malngLookup((lngCrc32 Xor Mid(strFile, i, 1)) And &HFF))
    Next i
    
'    lngFilePtr = CreateFile(strFile, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, 0, 0)
'
'    If lngFilePtr Then
'        lngFileLength = GetFileSize(lngFilePtr, 0)
'
'        If lngFileLength Then
'            Do
'                If ReadFile(lngFilePtr, abytBuffer(0), CRC_BUFFER_LENGTH, lngBufferLength, ByVal 0&) Then
'
'                    For i = 0 To lngBufferLength - 1
'                        lngCrc32 = (Int(lngCrc32 / 256) And &HFFFFFF) Xor (malngLookup((lngCrc32 Xor abytBuffer(i)) And &HFF))
'                    Next i
'
'                    lngBufferPtr = lngBufferPtr + lngBufferLength
'
'                    mobjConsole.CurrentX = 0
'                    mobjConsole.CWrite Fix(lngBufferPtr * (100 / lngFileLength)) & "%"
'
'                End If
'            Loop While lngBufferLength = CRC_BUFFER_LENGTH
'        End If
'
'        CloseHandle lngFilePtr
'    End If
    
    GetCrc32 = Not lngCrc32
    
End Function

Private Sub CreateCrc32Lookup()
    Static sblnDone As Boolean
    
    If sblnDone = False Then
        malngLookup(0) = &H0
        malngLookup(1) = &H77073096
        malngLookup(2) = &HEE0E612C
        malngLookup(3) = &H990951BA
        malngLookup(4) = &H76DC419
        malngLookup(5) = &H706AF48F
        malngLookup(6) = &HE963A535
        malngLookup(7) = &H9E6495A3
        malngLookup(8) = &HEDB8832
        malngLookup(9) = &H79DCB8A4
        malngLookup(10) = &HE0D5E91E
        malngLookup(11) = &H97D2D988
        malngLookup(12) = &H9B64C2B
        malngLookup(13) = &H7EB17CBD
        malngLookup(14) = &HE7B82D07
        malngLookup(15) = &H90BF1D91
        malngLookup(16) = &H1DB71064
        malngLookup(17) = &H6AB020F2
        malngLookup(18) = &HF3B97148
        malngLookup(19) = &H84BE41DE
        malngLookup(20) = &H1ADAD47D
        malngLookup(21) = &H6DDDE4EB
        malngLookup(22) = &HF4D4B551
        malngLookup(23) = &H83D385C7
        malngLookup(24) = &H136C9856
        malngLookup(25) = &H646BA8C0
        malngLookup(26) = &HFD62F97A
        malngLookup(27) = &H8A65C9EC
        malngLookup(28) = &H14015C4F
        malngLookup(29) = &H63066CD9
        malngLookup(30) = &HFA0F3D63
        malngLookup(31) = &H8D080DF5
        malngLookup(32) = &H3B6E20C8
        malngLookup(33) = &H4C69105E
        malngLookup(34) = &HD56041E4
        malngLookup(35) = &HA2677172
        malngLookup(36) = &H3C03E4D1
        malngLookup(37) = &H4B04D447
        malngLookup(38) = &HD20D85FD
        malngLookup(39) = &HA50AB56B
        malngLookup(40) = &H35B5A8FA
        malngLookup(41) = &H42B2986C
        malngLookup(42) = &HDBBBC9D6
        malngLookup(43) = &HACBCF940
        malngLookup(44) = &H32D86CE3
        malngLookup(45) = &H45DF5C75
        malngLookup(46) = &HDCD60DCF
        malngLookup(47) = &HABD13D59
        malngLookup(48) = &H26D930AC
        malngLookup(49) = &H51DE003A
        malngLookup(50) = &HC8D75180
        malngLookup(51) = &HBFD06116
        malngLookup(52) = &H21B4F4B5
        malngLookup(53) = &H56B3C423
        malngLookup(54) = &HCFBA9599
        malngLookup(55) = &HB8BDA50F
        malngLookup(56) = &H2802B89E
        malngLookup(57) = &H5F058808
        malngLookup(58) = &HC60CD9B2
        malngLookup(59) = &HB10BE924
        malngLookup(60) = &H2F6F7C87
        malngLookup(61) = &H58684C11
        malngLookup(62) = &HC1611DAB
        malngLookup(63) = &HB6662D3D
        malngLookup(64) = &H76DC4190
        malngLookup(65) = &H1DB7106
        malngLookup(66) = &H98D220BC
        malngLookup(67) = &HEFD5102A
        malngLookup(68) = &H71B18589
        malngLookup(69) = &H6B6B51F
        malngLookup(70) = &H9FBFE4A5
        malngLookup(71) = &HE8B8D433
        malngLookup(72) = &H7807C9A2
        malngLookup(73) = &HF00F934
        malngLookup(74) = &H9609A88E
        malngLookup(75) = &HE10E9818
        malngLookup(76) = &H7F6A0DBB
        malngLookup(77) = &H86D3D2D
        malngLookup(78) = &H91646C97
        malngLookup(79) = &HE6635C01
        malngLookup(80) = &H6B6B51F4
        malngLookup(81) = &H1C6C6162
        malngLookup(82) = &H856530D8
        malngLookup(83) = &HF262004E
        malngLookup(84) = &H6C0695ED
        malngLookup(85) = &H1B01A57B
        malngLookup(86) = &H8208F4C1
        malngLookup(87) = &HF50FC457
        malngLookup(88) = &H65B0D9C6
        malngLookup(89) = &H12B7E950
        malngLookup(90) = &H8BBEB8EA
        malngLookup(91) = &HFCB9887C
        malngLookup(92) = &H62DD1DDF
        malngLookup(93) = &H15DA2D49
        malngLookup(94) = &H8CD37CF3
        malngLookup(95) = &HFBD44C65
        malngLookup(96) = &H4DB26158
        malngLookup(97) = &H3AB551CE
        malngLookup(98) = &HA3BC0074
        malngLookup(99) = &HD4BB30E2
        malngLookup(100) = &H4ADFA541
        malngLookup(101) = &H3DD895D7
        malngLookup(102) = &HA4D1C46D
        malngLookup(103) = &HD3D6F4FB
        malngLookup(104) = &H4369E96A
        malngLookup(105) = &H346ED9FC
        malngLookup(106) = &HAD678846
        malngLookup(107) = &HDA60B8D0
        malngLookup(108) = &H44042D73
        malngLookup(109) = &H33031DE5
        malngLookup(110) = &HAA0A4C5F
        malngLookup(111) = &HDD0D7CC9
        malngLookup(112) = &H5005713C
        malngLookup(113) = &H270241AA
        malngLookup(114) = &HBE0B1010
        malngLookup(115) = &HC90C2086
        malngLookup(116) = &H5768B525
        malngLookup(117) = &H206F85B3
        malngLookup(118) = &HB966D409
        malngLookup(119) = &HCE61E49F
        malngLookup(120) = &H5EDEF90E
        malngLookup(121) = &H29D9C998
        malngLookup(122) = &HB0D09822
        malngLookup(123) = &HC7D7A8B4
        malngLookup(124) = &H59B33D17
        malngLookup(125) = &H2EB40D81
        malngLookup(126) = &HB7BD5C3B
        malngLookup(127) = &HC0BA6CAD
        malngLookup(128) = &HEDB88320
        malngLookup(129) = &H9ABFB3B6
        malngLookup(130) = &H3B6E20C
        malngLookup(131) = &H74B1D29A
        malngLookup(132) = &HEAD54739
        malngLookup(133) = &H9DD277AF
        malngLookup(134) = &H4DB2615
        malngLookup(135) = &H73DC1683
        malngLookup(136) = &HE3630B12
        malngLookup(137) = &H94643B84
        malngLookup(138) = &HD6D6A3E
        malngLookup(139) = &H7A6A5AA8
        malngLookup(140) = &HE40ECF0B
        malngLookup(141) = &H9309FF9D
        malngLookup(142) = &HA00AE27
        malngLookup(143) = &H7D079EB1
        malngLookup(144) = &HF00F9344
        malngLookup(145) = &H8708A3D2
        malngLookup(146) = &H1E01F268
        malngLookup(147) = &H6906C2FE
        malngLookup(148) = &HF762575D
        malngLookup(149) = &H806567CB
        malngLookup(150) = &H196C3671
        malngLookup(151) = &H6E6B06E7
        malngLookup(152) = &HFED41B76
        malngLookup(153) = &H89D32BE0
        malngLookup(154) = &H10DA7A5A
        malngLookup(155) = &H67DD4ACC
        malngLookup(156) = &HF9B9DF6F
        malngLookup(157) = &H8EBEEFF9
        malngLookup(158) = &H17B7BE43
        malngLookup(159) = &H60B08ED5
        malngLookup(160) = &HD6D6A3E8
        malngLookup(161) = &HA1D1937E
        malngLookup(162) = &H38D8C2C4
        malngLookup(163) = &H4FDFF252
        malngLookup(164) = &HD1BB67F1
        malngLookup(165) = &HA6BC5767
        malngLookup(166) = &H3FB506DD
        malngLookup(167) = &H48B2364B
        malngLookup(168) = &HD80D2BDA
        malngLookup(169) = &HAF0A1B4C
        malngLookup(170) = &H36034AF6
        malngLookup(171) = &H41047A60
        malngLookup(172) = &HDF60EFC3
        malngLookup(173) = &HA867DF55
        malngLookup(174) = &H316E8EEF
        malngLookup(175) = &H4669BE79
        malngLookup(176) = &HCB61B38C
        malngLookup(177) = &HBC66831A
        malngLookup(178) = &H256FD2A0
        malngLookup(179) = &H5268E236
        malngLookup(180) = &HCC0C7795
        malngLookup(181) = &HBB0B4703
        malngLookup(182) = &H220216B9
        malngLookup(183) = &H5505262F
        malngLookup(184) = &HC5BA3BBE
        malngLookup(185) = &HB2BD0B28
        malngLookup(186) = &H2BB45A92
        malngLookup(187) = &H5CB36A04
        malngLookup(188) = &HC2D7FFA7
        malngLookup(189) = &HB5D0CF31
        malngLookup(190) = &H2CD99E8B
        malngLookup(191) = &H5BDEAE1D
        malngLookup(192) = &H9B64C2B0
        malngLookup(193) = &HEC63F226
        malngLookup(194) = &H756AA39C
        malngLookup(195) = &H26D930A
        malngLookup(196) = &H9C0906A9
        malngLookup(197) = &HEB0E363F
        malngLookup(198) = &H72076785
        malngLookup(199) = &H5005713
        malngLookup(200) = &H95BF4A82
        malngLookup(201) = &HE2B87A14
        malngLookup(202) = &H7BB12BAE
        malngLookup(203) = &HCB61B38
        malngLookup(204) = &H92D28E9B
        malngLookup(205) = &HE5D5BE0D
        malngLookup(206) = &H7CDCEFB7
        malngLookup(207) = &HBDBDF21
        malngLookup(208) = &H86D3D2D4
        malngLookup(209) = &HF1D4E242
        malngLookup(210) = &H68DDB3F8
        malngLookup(211) = &H1FDA836E
        malngLookup(212) = &H81BE16CD
        malngLookup(213) = &HF6B9265B
        malngLookup(214) = &H6FB077E1
        malngLookup(215) = &H18B74777
        malngLookup(216) = &H88085AE6
        malngLookup(217) = &HFF0F6A70
        malngLookup(218) = &H66063BCA
        malngLookup(219) = &H11010B5C
        malngLookup(220) = &H8F659EFF
        malngLookup(221) = &HF862AE69
        malngLookup(222) = &H616BFFD3
        malngLookup(223) = &H166CCF45
        malngLookup(224) = &HA00AE278
        malngLookup(225) = &HD70DD2EE
        malngLookup(226) = &H4E048354
        malngLookup(227) = &H3903B3C2
        malngLookup(228) = &HA7672661
        malngLookup(229) = &HD06016F7
        malngLookup(230) = &H4969474D
        malngLookup(231) = &H3E6E77DB
        malngLookup(232) = &HAED16A4A
        malngLookup(233) = &HD9D65ADC
        malngLookup(234) = &H40DF0B66
        malngLookup(235) = &H37D83BF0
        malngLookup(236) = &HA9BCAE53
        malngLookup(237) = &HDEBB9EC5
        malngLookup(238) = &H47B2CF7F
        malngLookup(239) = &H30B5FFE9
        malngLookup(240) = &HBDBDF21C
        malngLookup(241) = &HCABAC28A
        malngLookup(242) = &H53B39330
        malngLookup(243) = &H24B4A3A6
        malngLookup(244) = &HBAD03605
        malngLookup(245) = &HCDD70693
        malngLookup(246) = &H54DE5729
        malngLookup(247) = &H23D967BF
        malngLookup(248) = &HB3667A2E
        malngLookup(249) = &HC4614AB8
        malngLookup(250) = &H5D681B02
        malngLookup(251) = &H2A6F2B94
        malngLookup(252) = &HB40BBE37
        malngLookup(253) = &HC30C8EA1
        malngLookup(254) = &H5A05DF1B
        malngLookup(255) = &H2D02EF8D
        
        sblnDone = True
    End If
    
End Sub

