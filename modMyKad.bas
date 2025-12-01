Attribute VB_Name = "modMyKad"
Option Explicit

' --- WINSCARD API DECLARATIONS ---
Private Declare Function SCardEstablishContext Lib "winscard.dll" (ByVal dwScope As Long, ByVal pvReserved1 As Long, ByVal pvReserved2 As Long, ByRef phContext As Long) As Long
Private Declare Function SCardReleaseContext Lib "winscard.dll" (ByVal hContext As Long) As Long
Private Declare Function SCardListReaders Lib "winscard.dll" Alias "SCardListReadersA" (ByVal hContext As Long, ByVal mszGroups As String, ByVal mszReaders As String, ByRef pcchReaders As Long) As Long
Private Declare Function SCardConnect Lib "winscard.dll" Alias "SCardConnectA" (ByVal hContext As Long, ByVal szReader As String, ByVal dwShareMode As Long, ByVal dwPreferredProtocols As Long, ByRef phCard As Long, ByRef pdwActiveProtocol As Long) As Long
Private Declare Function SCardDisconnect Lib "winscard.dll" (ByVal hCard As Long, ByVal dwDisposition As Long) As Long
Private Declare Function SCardTransmit Lib "winscard.dll" (ByVal hCard As Long, ByRef pioSendPci As Any, ByRef pbSendBuffer As Byte, ByVal cbSendLength As Long, ByRef pioRecvPci As Any, ByRef pbRecvBuffer As Byte, ByRef pcbRecvLength As Long) As Long

' --- CONSTANTS ---
Private Const SCARD_SCOPE_USER As Long = 0
Private Const SCARD_SHARE_SHARED As Long = 2
Private Const SCARD_PROTOCOL_T0 As Long = 1
Private Const SCARD_PROTOCOL_T1 As Long = 2
Private Const SCARD_LEAVE_CARD As Long = 0

Private Type SCARD_IO_REQUEST
    dwProtocol As Long
    cbPciLength As Long
End Type

' --- DATA STRUCTURE ---
Public Type MyKadData
    ICNo As String
    Name As String
    OldIC As String
    Gender As String
    DOB As String
    PlaceOfBirth As String
    Citizenship As String
    Race As String
    Religion As String
    Address1 As String
    Address2 As String
    Address3 As String
    Postcode As String
    City As String
    State As String
End Type

' --- MAIN FUNCTION ---
Public Function PerformMyKadScan(ByRef data As MyKadData) As Boolean
    Dim hContext As Long
    Dim hCard As Long
    Dim lRet As Long
    Dim sReaderList As String * 256
    Dim lReaderLen As Long
    Dim lActiveProtocol As Long
    Dim ioRequest As SCARD_IO_REQUEST
    
    ' 1. Establish Context
    lRet = SCardEstablishContext(SCARD_SCOPE_USER, 0, 0, hContext)
    If lRet <> 0 Then Exit Function
    
    ' 2. List Readers
    lReaderLen = 255
    lRet = SCardListReaders(hContext, vbNullString, sReaderList, lReaderLen)
    If lRet <> 0 Then
        SCardReleaseContext hContext
        Exit Function
    End If
    
    Dim sReader As String
    sReader = Left$(sReaderList, InStr(sReaderList, vbNullChar) - 1)
    
    ' 3. Connect
    lRet = SCardConnect(hContext, sReader, SCARD_SHARE_SHARED, SCARD_PROTOCOL_T0 Or SCARD_PROTOCOL_T1, hCard, lActiveProtocol)
    If lRet <> 0 Then
        SCardReleaseContext hContext
        Exit Function
    End If
    
    ioRequest.dwProtocol = lActiveProtocol
    ioRequest.cbPciLength = 8
    
    ' 4. Select JPN Applet
    If Not SendAPDU(hCard, ioRequest, "00A404000AA0000000744A504E0010") Then GoTo CleanUp
    
    ' 5. Get Response
    If Not SendAPDU(hCard, ioRequest, "00C0000005") Then GoTo CleanUp
    
    ' --- READ FILE 1: PERSONAL INFO (459 Bytes) ---
    Dim bJPN1() As Byte
    If ReadMyKadFile(hCard, ioRequest, 1, 459, bJPN1) Then
        data.Name = CleanString(ExtractString(bJPN1, 3, 40))
        data.ICNo = CleanString(ExtractString(bJPN1, 273, 13))
        data.Gender = CleanString(ExtractString(bJPN1, 286, 1))
        data.OldIC = CleanString(ExtractString(bJPN1, 287, 8))
        data.DOB = ParseDateString(ExtractBytes(bJPN1, 295, 4))
        data.PlaceOfBirth = CleanString(ExtractString(bJPN1, 299, 25))
        data.Citizenship = CleanString(ExtractString(bJPN1, 328, 18))
        data.Race = CleanString(ExtractString(bJPN1, 346, 25))
        data.Religion = CleanString(ExtractString(bJPN1, 371, 11))
    End If
    
    ' --- READ FILE 4: ADDRESS INFO (171 Bytes) ---
    Dim bJPN4() As Byte
    If ReadMyKadFile(hCard, ioRequest, 4, 171, bJPN4) Then
        data.Address1 = CleanString(ExtractString(bJPN4, 3, 30))
        data.Address2 = CleanString(ExtractString(bJPN4, 33, 30))
        data.Address3 = CleanString(ExtractString(bJPN4, 63, 30))
        data.Postcode = ParsePostcode(ExtractBytes(bJPN4, 93, 3))
        data.City = CleanString(ExtractString(bJPN4, 96, 25))
        data.State = CleanString(ExtractString(bJPN4, 121, 30))
    End If
    
    PerformMyKadScan = True

CleanUp:
    SCardDisconnect hCard, SCARD_LEAVE_CARD
    SCardReleaseContext hContext
End Function

Private Function ReadMyKadFile(hCard As Long, ioReq As SCARD_IO_REQUEST, iFileID As Integer, lFileLen As Long, ByRef bOut() As Byte) As Boolean
    Dim lOffset As Long
    Dim lChunk As Long
    Dim bCmd() As Byte
    Dim bRecv(260) As Byte
    Dim lRecvLen As Long
    Dim sHexFileID As String
    
    ReDim bOut(lFileLen) As Byte
    sHexFileID = Right$("0000" & Hex$(iFileID), 4)
    
    lOffset = 0
    Do While lOffset < lFileLen
        lChunk = 252
        If (lOffset + lChunk) > lFileLen Then lChunk = lFileLen - lOffset
        
        ' 1. Set Length (Little Endian Length)
        ' e.g. 252 -> FC 00
        Dim sLenLE As String
        sLenLE = Right$("0000" & Hex$(lChunk), 4)
        sLenLE = Right$(sLenLE, 2) & Left$(sLenLE, 2)
        
        SendAPDU hCard, ioReq, "C832000005080000" & sLenLE
        
        ' 2. Select File & Offset
        ' CC 00 00 00 08 [FileID] [01 00] [Offset LE] [Len LE]
        Dim sOffsetLE As String
        sOffsetLE = Right$("0000" & Hex$(lOffset), 4)
        sOffsetLE = Right$(sOffsetLE, 2) & Left$(sOffsetLE, 2)
        
        Dim sFileIDLE As String
        sFileIDLE = Right$("0000" & Hex$(iFileID), 4)
        sFileIDLE = Right$(sFileIDLE, 2) & Left$(sFileIDLE, 2)
        
        SendAPDU hCard, ioReq, "CC00000008" & sFileIDLE & "0100" & sOffsetLE & sLenLE
        
        ' 3. Get Data
        bCmd = HexToByte("CC060000" & Right$("00" & Hex$(lChunk), 2))
        lRecvLen = 260
        SCardTransmit hCard, ioReq, bCmd(0), UBound(bCmd) + 1, ioReq, bRecv(0), lRecvLen
        
        ' Copy Data
        Dim i As Long
        For i = 0 To lChunk - 1
            If (lOffset + i) < UBound(bOut) Then
                bOut(lOffset + i) = bRecv(i)
            End If
        Next i
        lOffset = lOffset + lChunk
    Loop
    
    ReadMyKadFile = True
End Function

' --- HELPERS ---
Private Function SendAPDU(hCard As Long, ioReq As SCARD_IO_REQUEST, sHex As String) As Boolean
    Dim bCmd() As Byte
    Dim bRecv(255) As Byte
    Dim lRecvLen As Long
    
    bCmd = HexToByte(sHex)
    lRecvLen = 255
    If SCardTransmit(hCard, ioReq, bCmd(0), UBound(bCmd) + 1, ioReq, bRecv(0), lRecvLen) = 0 Then
        SendAPDU = True
    End If
End Function

Private Function ExtractString(bArr() As Byte, lStart As Long, lLen As Long) As String
    Dim i As Long, s As String
    s = ""
    For i = 0 To lLen - 1
        If bArr(lStart + i) >= 32 And bArr(lStart + i) <= 126 Then
            s = s & Chr(bArr(lStart + i))
        End If
    Next i
    ExtractString = s
End Function

Private Function ExtractBytes(bArr() As Byte, lStart As Long, lLen As Long) As Byte()
    Dim bOut() As Byte
    ReDim bOut(lLen - 1)
    Dim i As Long
    For i = 0 To lLen - 1
        bOut(i) = bArr(lStart + i)
    Next i
    ExtractBytes = bOut
End Function

Private Function CleanString(s As String) As String
    CleanString = Trim(Replace(s, vbNullChar, ""))
End Function

Private Function ParseDateString(bData() As Byte) As String
    ' YYYY-MM-DD
    On Error Resume Next
    ParseDateString = Hex$(bData(0)) & Hex$(bData(1)) & "-" & Hex$(bData(2)) & "-" & Hex$(bData(3))
End Function

Private Function ParsePostcode(bData() As Byte) As String
    ParsePostcode = Hex$(bData(0)) & Hex$(bData(1)) & Right$(Hex$(bData(2)), 1)
End Function

Private Function HexToByte(ByVal sHex As String) As Byte()
    Dim i As Long
    Dim b() As Byte
    ReDim b(Len(sHex) / 2 - 1)
    For i = 0 To UBound(b)
        b(i) = CByte("&H" & Mid(sHex, i * 2 + 1, 2))
    Next i
    HexToByte = b
End Function
