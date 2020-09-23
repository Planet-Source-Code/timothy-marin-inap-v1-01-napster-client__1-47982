Attribute VB_Name = "Mp3Info"
Type HeaderInfo
    MpegLayer As String
    Frequency As String
    Bitrate As String
    Mode As String
    MpegVersion As String
    Emphasis As String
    PlayTime As String
    Duration As String
    mFileSize As String
End Type
Public MP3HeaderInfo As HeaderInfo

Public Function ReadHeader(ByVal FileName As String)
    On Error GoTo ErrHand
    
    Dim ByteArray(4) As Byte
    Dim XingH As String * 4
    Dim FIO As Long
    Dim i As Long
    Dim z As Long
    Dim X As Byte
    Dim HeadStart As Long
    Dim Frames As Long
    Dim Bin As String
    Dim Temp As Variant
    Dim Brate As Variant
    Dim Freq As Variant
         
    'tables
    Dim VersionLayer(3) As String
    VersionLayer(0) = 0
    VersionLayer(1) = 3
    VersionLayer(2) = 2
    VersionLayer(3) = 1
    
    Dim SMode(3) As String
    SMode(0) = "stereo"
    SMode(1) = "joint stereo"
    SMode(2) = "dual channel"
    SMode(3) = "single channel"
    
    
    FIO = FreeFile
    
    'read the header
    Open FileName For Binary Access Read As FIO
        If LOF(FIO) < 256 Then
            Close FIO
            Exit Function
        End If
        
        '''''start check startposition for header''''''''''''
        '''''if start position <>1 then id3v2 tag exists'''''
        For i = 1 To LOF(FIO)           'check the whole file for the header
            Get #FIO, i, X
            If X = 255 Then             'header always start with 255 followed by 250 or 251
                Get #FIO, i + 1, X
                If X > 249 And X < 252 Then
                    HeadStart = i       'set header start position
                    Exit For
                End If
            End If
        Next i
        
        'no header start position was found
        If HeadStart = 0 Then
            Exit Function
        End If
        '''end check start position for header'''''''''''''
    
        ''start check for XingHeader'''
        Get #FIO, HeadStart + 36, XingH
        If XingH = "Xing" Then
            VBR = True
            For z = 1 To 4 '
                Get #1, HeadStart + 43 + z, ByteArray(z)  'get framelength to array
            Next z
            Frames = BinToDec(ByteToBit(ByteArray))   'calculate # of frames
        Else
            VBR = False
        End If
        '''end check for XingHeader
    
        '''start extract the first 4 bytes (32 bits) to an array
        For z = 1 To 4 '
            Get #FIO, HeadStart + z - 1, ByteArray(z)
        Next z
        '''stop extract the first 4 bytes (32 bits) to an array
    Close FIO
    
    'header string
    Bin = ByteToBit(ByteArray)
    
    
    'get mpegversion from table
    MP3HeaderInfo.MpegVersion = VersionLayer(BinToDec(Mid(Bin, 12, 2)))
    ''get layer from table
    MP3HeaderInfo.MpegLayer = VersionLayer(BinToDec(Mid(Bin, 14, 2)))
    ''get mode from table
    MP3HeaderInfo.Mode = SMode(BinToDec(Mid(Bin, 25, 2)))
    
    'look for version to create right table
    Select Case MP3HeaderInfo.MpegVersion
        Case 1
            'for version 1
            Freq = Array(44100, 48000, 32000)
        Case 2 Or 25
            'for version 2 or 2.5
            Freq = Array(22050, 24000, 16000)
        Case Else
            MP3HeaderInfo.Frequency = 0
            Exit Function
    End Select
    
    'look for frequency in table
    MP3HeaderInfo.Frequency = Freq(BinToDec(Mid(Bin, 21, 2)))
    
    If VBR = True Then
        'define to calculate correct bitrate
        Temp = Array(, 12, 144, 144)
        MP3HeaderInfo.Bitrate = (FileLen(FileName) * MP3HeaderInfo.Frequency) / (Int(Frames)) / 1000 / Temp(MP3HeaderInfo.MpegLayer)
    Else
        'look for the right bitrate table
        Select Case Val(MP3HeaderInfo.MpegVersion & MP3HeaderInfo.MpegLayer)
            Case 11
                'Version 1, Layer 1
                Brate = Array(0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448)
            Case 12
                'V1 L1
                Brate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384)
            Case 13
                'V1 L3
                Brate = Array(0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320)
            Case 21 Or 251
                'V2 L1 and 'V2.5 L1
                Brate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 144, 160, 176, 192, 224, 256)
            Case 22 Or 252 Or 23 Or 253
                'V2 L2 and 'V2.5 L2 etc...
                Brate = Array(0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160)
            Case Else
                'if variable bitrate
                MP3HeaderInfo.Bitrate = 1
            Exit Function
        End Select
        
        MP3HeaderInfo.Bitrate = Brate(BinToDec(Mid(Bin, 17, 4)))
    End If
    
    'if there is a decimal place, then parse it off
    If InStr(1, MP3HeaderInfo.Bitrate, ".") Then
        MP3HeaderInfo.Bitrate = Left(MP3HeaderInfo.Bitrate, InStr(1, MP3HeaderInfo.Bitrate, ".") - 1)
    End If
    
    'calculate duration
    MP3HeaderInfo.Duration = Int((FileLen(FileName) * 8) / MP3HeaderInfo.Bitrate / 1000)
    MP3HeaderInfo.PlayTime = MP3HeaderInfo.Duration
    MP3HeaderInfo.Duration = MP3HeaderInfo.Duration \ 60 & ":" & Format(MP3HeaderInfo.Duration - (MP3HeaderInfo.Duration \ 60) * 60, "0#")
    
    
    Exit Function
    
ErrHand:
    'Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3 (" & FileName & ")"
    Close FIO
    
End Function

Private Function BinToDec(BinValue As String) As Long
    On Error GoTo ErrHand
    
    Dim i As Long
    BinToDec = 0
    For i = 1 To Len(BinValue)
        If Mid(BinValue, i, 1) = 1 Then
            BinToDec = BinToDec + 2 ^ (Len(BinValue) - i)
        End If
    Next i
    
    Exit Function
    
ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3"

End Function

Private Function ByteToBit(ByteArray) As String
    On Error GoTo ErrHand

    Dim z As Integer
    Dim i As Integer
    'convert 4*1 byte array to 4*8 bits'''''
    ByteToBit = ""
    For z = 1 To 4
        For i = 7 To 0 Step -1
            If Int(ByteArray(z) / (2 ^ i)) = 1 Then
                ByteToBit = ByteToBit & "1"
                ByteArray(z) = ByteArray(z) - (2 ^ i)
            ElseIf ByteToBit <> "" Then
                ByteToBit = ByteToBit & "0"
            End If
        Next
    Next z
    
    Exit Function
    
ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3"

End Function
