Attribute VB_Name = "modTestTranscribe"
Option Explicit

' Test pipeline dong bo: doc WAV/MP3/MP4/M4A -> resample 16kHz mono ->
' transcribe bang whisper.cpp. Can modDllHelper.Auto_Open da chay truoc.

' ------------------- Declare -------------------
' Tren Win64, stdcall va cdecl la CUNG 1 calling convention o tang ABI
' (khac Win32) - Declare PtrSafe binh thuong la du, khong can phan biet.

Private Declare PtrSafe Function ResampleToWhisperFormat Lib "WavParser.dll" ( _
    ByVal lpFileName As LongPtr, _
    ByVal pOutBuffer As LongPtr, _
    ByVal outBufferCapacity As Long, _
    ByRef pOutSampleCount As Long) As Long

Private Declare PtrSafe Function TranscribeWavFile Lib "transcribe.dll" ( _
    ByVal modelPath As LongPtr, _
    ByVal pcmSamples As LongPtr, _
    ByVal SampleCount As Long, _
    ByVal languageCode As LongPtr, _
    ByVal OutBuf As LongPtr, _
    ByVal outBufLen As Long) As Long

' Chi dung de debug - xuat lai buffer da resample ra file .wav de nghe thu
Private Declare PtrSafe Function WriteWavFile Lib "WavParser.dll" ( _
    ByVal lpFileName As LongPtr, _
    ByVal pSamples As LongPtr, _
    ByVal SampleCount As Long, _
    ByVal sampleRate As Long) As Long

' transcribe.dll tra chu tieng Viet dang UTF-8 - StrConv(...,vbUnicode) hieu
' sai thanh ANSI se ra chu rac (mojibake), can giai ma dung bang ham nay
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long) As Long

' ------------------- Const -------------------
Private Const OUTPUT_BUFFER_SIZE As Long = 8192 ' du cho vai cau transcribe
Private Const CP_UTF8 As Long = 65001

' ------------------- Wrapper -------------------

Private Function UTF8BytesToString(ByRef bytes() As Byte, ByVal byteLen As Long) As String
    If byteLen <= 0 Then
        UTF8BytesToString = ""
        Exit Function
    End If

    Dim charCount As Long
    charCount = MultiByteToWideChar(CP_UTF8, 0, VarPtr(bytes(0)), byteLen, 0, 0)
    If charCount <= 0 Then
        UTF8BytesToString = ""
        Exit Function
    End If

    Dim Buf As String
    Buf = String$(charCount, vbNullChar)
    MultiByteToWideChar CP_UTF8, 0, VarPtr(bytes(0)), byteLen, StrPtr(Buf), charCount
    UTF8BytesToString = Buf
End Function

Public Function DoTranscribe(ByVal WavPath As String, ByVal Language As String) As String
    ' Loi tra ve dang chuoi bat dau "ERROR:" kem ma loi de de doc
    Dim SampleCount As Long
    Dim Buf() As Integer
    Dim OutBuf() As Byte
    Dim ResultLen As Long
    Dim ModelPathBytes() As Byte
    Dim LangBytes() As Byte

    If modDllHelper.WavParserHandle = 0 Or modDllHelper.TranscribeHandle = 0 Then
        DoTranscribe = "ERROR: DLL chua duoc load - kiem tra Auto_Open da chay chua"
        Exit Function
    End If

    If Not CreateObject("Scripting.FileSystemObject").FileExists(modDllHelper.modelPath) Then
        DoTranscribe = "ERROR: Khong tim thay model tai " & modDllHelper.modelPath
        Exit Function
    End If

    ' Hoi truoc so luong sample can cap phat, roi doc du lieu PCM that su
    ResampleToWhisperFormat StrPtr(WavPath), 0, 0, SampleCount
    If SampleCount <= 0 Then
        DoTranscribe = "ERROR: Khong doc duoc file WAV hoac file rong"
        Exit Function
    End If

    ReDim Buf(SampleCount - 1)
    If ResampleToWhisperFormat(StrPtr(WavPath), VarPtr(Buf(0)), SampleCount, SampleCount) = 0 Then
        DoTranscribe = "ERROR: Resample that bai"
        Exit Function
    End If

    ' Debug: xuat lai buffer da resample de nghe thu, xac nhan loi (neu co)
    ' khong phai o buoc resample truoc khi nghi cho whisper
    WriteWavFile StrPtr(ThisWorkbook.path & "\debug_resampled.wav"), VarPtr(Buf(0)), SampleCount, 16000

    ' Cac bien byte() phai song het qua luc goi ham DLL - khong dung ham
    ' phu tra ve VarPtr roi ket thuc ham, bien cuc bo se bi giai phong
    ' ngay khi ham do return, con tro tra ve tro vao vung nho da giai
    ' phong (loi rat kho phat hien)
    ModelPathBytes = StrConv(modDllHelper.modelPath & Chr(0), vbFromUnicode)
    LangBytes = StrConv(Language & Chr(0), vbFromUnicode)
    ReDim OutBuf(OUTPUT_BUFFER_SIZE - 1)

    ResultLen = TranscribeWavFile( _
        VarPtr(ModelPathBytes(0)), VarPtr(Buf(0)), SampleCount, _
        VarPtr(LangBytes(0)), VarPtr(OutBuf(0)), OUTPUT_BUFFER_SIZE)

    If ResultLen < 0 Then
        DoTranscribe = "ERROR: TranscribeWavFile tra ve ma loi " & ResultLen & _
            " (-1=load model loi, -2=tao context loi, -3=loi transcribe, " & _
            "khac=buffer qua nho, can " & Abs(ResultLen) & " byte)"
        Exit Function
    End If

    DoTranscribe = UTF8BytesToString(OutBuf, ResultLen)
End Function

' ------------------- Test -------------------

Sub TestTranscribeFromFile()
    Dim WavPath As String, Result As String

    WavPath = Application.GetOpenFilename("Tap tin WAV (*.wav), *.wav", , "Chon file ghi am de transcribe")
    If WavPath = "False" Then Exit Sub

    Application.StatusBar = "Dang transcribe, vui long doi..."
    Result = DoTranscribe(WavPath, "vi")
    Application.StatusBar = False

    If Left(Result, 6) = "ERROR:" Then
        MsgBox Result, vbCritical, "Transcribe that bai"
    Else
        MsgBox "Ket qua:" & Chr(13) & Chr(13) & Result, vbInformation, "Transcribe thanh cong"
    End If
End Sub

Sub TestTranscribeToSheet()
    ' Ghi ket qua vao ActiveCell thay vi MsgBox - de doc/copy chu co dau
    ' de dang hon, khong phu thuoc font hien thi cua hop thoai
    Dim WavPath As Variant, Result As String, Target As Range

    WavPath = Application.GetOpenFilename( _
        "Audio/Video (*.wav;*.mp3;*.mp4;*.m4a),*.wav;*.mp3;*.mp4;*.m4a", _
        , "Chon file WAV, MP3, MP4 hoac M4A")
    If WavPath = False Then Exit Sub

    Set Target = ActiveCell

    Application.ScreenUpdating = False
    Application.StatusBar = "Dang transcribe, vui long doi..."
    Result = DoTranscribe(CStr(WavPath), "vi")
    Application.StatusBar = False
    Application.ScreenUpdating = True

    Target.Value = Result
    Target.WrapText = True
    Target.EntireRow.AutoFit

    If Left$(Result, 6) = "ERROR:" Then
        MsgBox "Co loi - xem chi tiet trong o " & Target.Address, vbCritical
    Else
        MsgBox "Da ghi ket qua vao o " & Target.Address, vbInformation
    End If
End Sub
