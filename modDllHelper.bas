Attribute VB_Name = "modDllHelper"
Option Explicit

' Load/unload WavParser.dll + transcribe.dll tu thu muc chua file Excel
' (Unicode-safe). Tinh san duong dan model GGML dung chung cho cac module
' transcribe khac. Goi Auto_Open truoc khi dung bat ky ham transcribe nao.

' ------------------- Declare -------------------
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As LongPtr) As LongPtr
    Public Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
    Public Declare PtrSafe Function SetCurrentDirectoryW Lib "kernel32" (ByVal lpPathName As LongPtr) As Long
    Public Declare PtrSafe Function SetDllDirectoryW Lib "kernel32" (ByVal lpPathName As LongPtr) As Long

    Public WavParserHandle As LongPtr
    Public TranscribeHandle As LongPtr
#Else
    Public Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long
    Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
    Public Declare Function SetCurrentDirectoryW Lib "kernel32" (ByVal lpPathName As Long) As Long
    Public Declare Function SetDllDirectoryW Lib "kernel32" (ByVal lpPathName As Long) As Long

    Public WavParserHandle As Long
    Public TranscribeHandle As Long
#End If

' ------------------- Const -------------------
Private Const WAVPARSER_DLL As String = "WavParser.dll"
Private Const TRANSCRIBE_DLL As String = "transcribe.dll"
Private Const MODEL_FILE As String = "ggml-small.bin"

Public modelPath As String ' tinh san 1 lan luc Auto_Open, dung lai o cac module khac

' ------------------- Wrapper -------------------

Public Sub ChDirW(ByVal s As String)
    If SetCurrentDirectoryW(StrPtr(s)) = 0 Then Err.Raise 76
End Sub

Public Function LoadDllFromWorkbookFolder(ByVal DllName As String) As LongPtr
    Dim FullPath As String, DllDir As String
    FullPath = ThisWorkbook.path & "\" & DllName
    DllDir = ThisWorkbook.path

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(FullPath) Then
        MsgBox "Khong tim thay DLL: " & FullPath, vbCritical
        LoadDllFromWorkbookFolder = 0
        Exit Function
    End If

    ' Bao Windows tim dependency (DLL phu) trong thu muc Excel file, roi reset lai
    SetDllDirectoryW StrPtr(DllDir)
    LoadDllFromWorkbookFolder = LoadLibraryW(StrPtr(FullPath))
    SetDllDirectoryW 0

    If LoadDllFromWorkbookFolder = 0 Then
        MsgBox "Khong the load DLL: " & FullPath & Chr(13) & "Error: " & Err.LastDllError, vbCritical
    End If
End Function

Public Sub FreeDllHandle(ByVal H As LongPtr)
    If H <> 0 Then FreeLibrary H
End Sub

Private Function CheckModelFileExists(ByVal FullPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    CheckModelFileExists = fso.FileExists(FullPath)
End Function

' ------------------- Auto_Open / Auto_Close -------------------

Sub Auto_Open()
    WavParserHandle = LoadDllFromWorkbookFolder(WAVPARSER_DLL)
    TranscribeHandle = LoadDllFromWorkbookFolder(TRANSCRIBE_DLL)

    If WavParserHandle <> 0 Then Debug.Print WAVPARSER_DLL & " loaded: 0x" & Hex$(WavParserHandle)
    If TranscribeHandle <> 0 Then Debug.Print TRANSCRIBE_DLL & " loaded: 0x" & Hex$(TranscribeHandle)

    modelPath = ThisWorkbook.path & "\" & MODEL_FILE
    If Not CheckModelFileExists(modelPath) Then
        MsgBox "Khong tim thay file model: " & modelPath & Chr(13) & _
               "Tai ve tu huggingface.co/ggerganov/whisper.cpp va dat cung thu muc file Excel nay.", _
               vbExclamation
    End If
End Sub

Sub Auto_Close()
    If WavParserHandle <> 0 Then
        FreeDllHandle WavParserHandle
        WavParserHandle = 0
        Debug.Print WAVPARSER_DLL & " unloaded"
    End If
    If TranscribeHandle <> 0 Then
        FreeDllHandle TranscribeHandle
        TranscribeHandle = 0
        Debug.Print TRANSCRIBE_DLL & " unloaded"
    End If
End Sub
