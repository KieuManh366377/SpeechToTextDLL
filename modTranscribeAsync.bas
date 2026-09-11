Attribute VB_Name = "modTranscribeAsync"
Option Explicit

' Ban ASYNC cua transcribe - danh cho file ghi am dai, khong "dong" Excel
' trong luc cho, hien progress % len StatusBar. Can modDllHelper.Auto_Open
' da chay truoc.

' ------------------- Declare -------------------

Private Declare PtrSafe Function ResampleToWhisperFormat Lib "WavParser.dll" ( _
    ByVal lpFileName As LongPtr, _
    ByVal pOutBuffer As LongPtr, _
    ByVal outBufferCapacity As Long, _
    ByRef pOutSampleCount As Long) As Long

Private Declare PtrSafe Function StartTranscribeAsync Lib "transcribe.dll" ( _
    ByVal modelPath As LongPtr, _
    ByVal pcmSamples As LongPtr, _
    ByVal SampleCount As Long, _
    ByVal languageCode As LongPtr) As Long

Private Declare PtrSafe Function GetTranscribeProgress Lib "transcribe.dll" () As Long
Private Declare PtrSafe Function IsTranscribeDone Lib "transcribe.dll" () As Long

Private Declare PtrSafe Function GetTranscribeResult Lib "transcribe.dll" ( _
    ByVal OutBuf As LongPtr, _
    ByVal outBufLen As Long) As Long

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long) As Long

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ------------------- Const -------------------
Private Const OUTPUT_BUFFER_SIZE As Long = 65536 ' lon hon ban sync, file dai ra nhieu chu hon
Private Const CP_UTF8 As Long = 65001
Private Const POLL_INTERVAL_MS As Long = 300 ' khoang cach giua cac lan hoi tien do

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

Public Function DoTranscribeAsync(ByVal WavPath As String, ByVal Language As String, _
    Optional ByVal ShowProgressOnStatusBar As Boolean = True) As String
    ' Giong DoTranscribe (ban sync) nhung khong lam Excel bi "dong" - vong
    ' lap ben duoi nhuong CPU bang DoEvents, Excel van tuong tac duoc binh
    ' thuong trong luc cho ket qua.
    Dim SampleCount As Long
    Dim Buf() As Integer
    Dim ModelPathBytes() As Byte
    Dim LangBytes() As Byte
    Dim StartResult As Long
    Dim Progress As Long

    If modDllHelper.WavParserHandle = 0 Or modDllHelper.TranscribeHandle = 0 Then
        DoTranscribeAsync = "ERROR: DLL chua duoc load - kiem tra Auto_Open da chay chua"
        Exit Function
    End If

    If Not CreateObject("Scripting.FileSystemObject").FileExists(modDllHelper.modelPath) Then
        DoTranscribeAsync = "ERROR: Khong tim thay model tai " & modDllHelper.modelPath
        Exit Function
    End If

    ' Resample - buoc nay van chay dong bo (thuong vai giay, khong dang ke
    ' so voi thoi gian transcribe)
    ResampleToWhisperFormat StrPtr(WavPath), 0, 0, SampleCount
    If SampleCount <= 0 Then
        DoTranscribeAsync = "ERROR: Khong doc duoc file WAV hoac file rong"
        Exit Function
    End If
    ReDim Buf(SampleCount - 1)
    If ResampleToWhisperFormat(StrPtr(WavPath), VarPtr(Buf(0)), SampleCount, SampleCount) = 0 Then
        DoTranscribeAsync = "ERROR: Resample that bai"
        Exit Function
    End If

    ModelPathBytes = StrConv(modDllHelper.modelPath & Chr(0), vbFromUnicode)
    LangBytes = StrConv(Language & Chr(0), vbFromUnicode)

    StartResult = StartTranscribeAsync( _
        VarPtr(ModelPathBytes(0)), VarPtr(Buf(0)), SampleCount, VarPtr(LangBytes(0)))

    If StartResult = -10 Then
        DoTranscribeAsync = "ERROR: Dang co 1 job transcribe khac chua xong - doi no xong truoc"
        Exit Function
    ElseIf StartResult <> 0 Then
        DoTranscribeAsync = "ERROR: Khong bat dau duoc job, ma loi " & StartResult
        Exit Function
    End If

    Do While IsTranscribeDone() = 0
        Progress = GetTranscribeProgress()
        If ShowProgressOnStatusBar Then Application.StatusBar = "Dang transcribe: " & Progress & "% ..."
        DoEvents
        Sleep POLL_INTERVAL_MS
    Loop

    If ShowProgressOnStatusBar Then Application.StatusBar = False

    Dim OutBuf() As Byte
    Dim ResultLen As Long
    ReDim OutBuf(OUTPUT_BUFFER_SIZE - 1)
    ResultLen = GetTranscribeResult(VarPtr(OutBuf(0)), OUTPUT_BUFFER_SIZE)

    If ResultLen < 0 Then
        DoTranscribeAsync = "ERROR: GetTranscribeResult tra ve ma loi " & ResultLen & _
            " (-1=load model loi, -2=tao context loi, -3=loi transcribe, " & _
            "khac=buffer qua nho, can " & Abs(ResultLen) & " byte - tang OUTPUT_BUFFER_SIZE)"
        Exit Function
    End If

    DoTranscribeAsync = UTF8BytesToString(OutBuf, ResultLen)
End Function

' ------------------- Test -------------------

Sub TestTranscribeAsyncToSheet()
    Dim WavPath As String, Result As String

    WavPath = Application.GetOpenFilename( _
        "Audio/Video (*.wav;*.mp3;*.mp4;*.m4a),*.wav;*.mp3;*.mp4;*.m4a", _
        , "Chon file WAV, MP3, MP4 hoac M4A")
    If WavPath = "False" Then Exit Sub

    Result = DoTranscribeAsync(WavPath, "vi")
    ActiveCell.Value = Result

    If Left(Result, 6) = "ERROR:" Then
        MsgBox "Co loi - xem chi tiet trong o " & ActiveCell.Address, vbCritical
    Else
        MsgBox "Da ghi ket qua vao o " & ActiveCell.Address, vbInformation
    End If
End Sub

' Ham goi truc tiep tren o tinh: =TRANSCRIBE(A1,"vi")
' Luu y: goi DoEvents/Sleep ben trong 1 UDF la cach lam khong chinh thong -
' Excel khong dam bao hanh vi UDF trong luc tinh toan cong thuc khi co
' tuong tac nguoi dung xen vao giua chung. Da dung on dinh trong test noi
' bo, nhung can tu kiem tra ky voi khoi luong file/tan suat goi thuc te.
Public Function TRANSCRIBE(ByVal WavPath As String, ByVal Lang As String) As String
    On Error GoTo ErrHandle

    If Len(Trim$(WavPath)) = 0 Then
        TRANSCRIBE = ""
        Exit Function
    End If

    If Dir$(WavPath) = "" Then
        TRANSCRIBE = "ERROR: File khong ton tai"
        Exit Function
    End If

    TRANSCRIBE = DoTranscribeAsync(WavPath, Lang)
    Exit Function

ErrHandle:
    TRANSCRIBE = "ERROR: " & Err.Description
End Function
