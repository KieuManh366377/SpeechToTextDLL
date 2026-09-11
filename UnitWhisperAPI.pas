//---------------------------------------------------------------------------
unit UnitWhisperAPI;
// Lop bridge goi 2 DLL doc lap: WavParser.dll (C++ Builder, resample am
// thanh WAV/MP3/MP4/M4A ve 16kHz mono PCM) va transcribe.dll (Go/cgo,
// wrap whisper.cpp, chuyen PCM thanh van ban). Khong phu thuoc VCL/Form -
// co the tai su dung nguyen ven cho InputInExcel (add-in) sau nay.
//---------------------------------------------------------------------------

interface

uses
  System.SysUtils;

//---------------------------------------------------------------------------
// Khai bao ham ngoai - WavParser.dll (dung __stdcall dung nhu khai bao
// goc trong WavParser.h ben C++ Builder)
//---------------------------------------------------------------------------
function ResampleToWhisperFormat(lpFileName: PWideChar; pOutBuffer: PSmallInt;
  outBufferCapacity: Cardinal; var pOutSampleCount: Cardinal): LongBool; stdcall;
  external 'WavParser.dll';

function ValidateWavForWhisper(lpFileName: PWideChar): LongBool; stdcall;
  external 'WavParser.dll';

//---------------------------------------------------------------------------
// Khai bao ham ngoai - transcribe.dll (Go, xuat qua //export - tren Win64
// stdcall/cdecl cho ra cung 1 ABI, khai bao cdecl cho dung ban chat cua
// Go c-shared build)
//---------------------------------------------------------------------------
function TranscribeWavFile(modelPath: PAnsiChar; pcmSamples: PSmallInt;
  sampleCount: Integer; languageCode: PAnsiChar; outBuf: PAnsiChar;
  outBufLen: Integer): Integer; cdecl; external 'transcribe.dll';

// TranscribeWavFileEx: giong TranscribeWavFile, them tham so initialPrompt
// de goi y truoc tu/ten rieng cho whisper - tang do chinh xac voi audio co
// tap am/phat am khong chuan. Truyen '' neu khong can goi y.
function TranscribeWavFileEx(modelPath: PAnsiChar; pcmSamples: PSmallInt;
  sampleCount: Integer; languageCode: PAnsiChar; initialPrompt: PAnsiChar;
  outBuf: PAnsiChar; outBufLen: Integer): Integer; cdecl; external 'transcribe.dll';

procedure ReleaseModel(modelPath: PAnsiChar); cdecl; external 'transcribe.dll';

//---------------------------------------------------------------------------
// Ham tien ich muc cao: nhan duong dan file am thanh bat ky (WAV/MP3/MP4/
// M4A), tu goi WavParser.dll de resample roi goi transcribe.dll de chuyen
// thanh van ban. Dung duong DONG BO (TranscribeWavFileEx) - phu hop file
// test ngan; file dai nen chuyen sang StartTranscribeAsync + polling ben
// transcribe.dll de khong treo UI (chua lam trong ban demo nay).
//
// AInitialPrompt (tuy chon, mac dinh rong): goi y truoc cho whisper -
// xem giai thich trong main.go/TranscribeWavFileEx.
//---------------------------------------------------------------------------
function TranscribeAudioFile(const AFileName, AModelPath, ALanguage: string;
  out AResultText: string; out AErrorMsg: string;
  const AInitialPrompt: string = ''): Boolean;

implementation

const
  // Cac ma loi tra ve tu transcribe.dll - doi chieu voi chu thich trong
  // main.go (ban da vá) de hien thi thong bao de hieu hon cho nguoi dung
  ERR_MODEL_LOAD_FAILED = -1;
  ERR_CONTEXT_FAILED = -2;
  ERR_PROCESS_FAILED = -3;
  ERR_TRANSCRIBE_BUSY = -11;
  ERR_INVALID_LANGUAGE = -12;
  ERR_OUTBUF_NULL = -13;
  ERR_MODEL_PATH_NULL = -20;
  ERR_INVALID_PCM = -21;
  ERR_INTERNAL_PANIC = -99;

function MoTaMaLoiTranscribe(AErrCode: Integer): string;
begin
  case AErrCode of
    ERR_MODEL_LOAD_FAILED: Result := 'Khong load duoc model (kiem tra duong dan ggml-small.bin)';
    ERR_CONTEXT_FAILED: Result := 'Loi tao context xu ly whisper';
    ERR_PROCESS_FAILED: Result := 'Loi trong qua trinh transcribe';
    ERR_TRANSCRIBE_BUSY: Result := 'Dang co 1 lan transcribe khac chay, thu lai sau';
    ERR_INVALID_LANGUAGE: Result := 'Ma ngon ngu khong hop le';
    ERR_OUTBUF_NULL: Result := 'Loi noi bo: buffer ket qua NULL';
    ERR_MODEL_PATH_NULL: Result := 'Duong dan model la NULL';
    ERR_INVALID_PCM: Result := 'Du lieu PCM khong hop le (rong hoac sample count <= 0)';
    ERR_INTERNAL_PANIC: Result := 'Loi noi bo khong luong truoc duoc trong transcribe.dll (da chan panic, xem lai input truyen vao)';
  else
    if AErrCode < 0 then
      Result := Format('Buffer ket qua qua nho, can toi thieu %d byte', [-AErrCode - 1])
    else
      Result := Format('Ma loi khong xac dinh: %d', [AErrCode]);
  end;
end;

function TranscribeAudioFile(const AFileName, AModelPath, ALanguage: string;
  out AResultText: string; out AErrorMsg: string;
  const AInitialPrompt: string = ''): Boolean;
var
  FileNameW: PWideChar;
  SampleCount: Cardinal;
  PcmBuffer: TArray<SmallInt>;
  ModelPathUtf8, LanguageUtf8, InitialPromptUtf8: UTF8String;
  PromptParam: PAnsiChar;
  OutBufLen: Integer;
  OutBuf: TArray<AnsiChar>;
  RawResult: RawByteString;
  RetCode: Integer;
begin
  Result := False;
  AResultText := '';
  AErrorMsg := '';

  if not FileExists(AFileName) then
  begin
    AErrorMsg := 'Khong tim thay file: ' + AFileName;
    Exit;
  end;
  if not FileExists(AModelPath) then
  begin
    AErrorMsg := 'Khong tim thay model: ' + AModelPath;
    Exit;
  end;

  FileNameW := PWideChar(AFileName);

  // --- Buoc 1: goi voi pOutBuffer = nil de hoi truoc so luong sample can
  // cap phat (dung 2-pass API theo dung thiet ke cua WavParser.dll) ---
  SampleCount := 0;
  if not ResampleToWhisperFormat(FileNameW, nil, 0, SampleCount) then
  begin
    AErrorMsg := 'WavParser.dll: khong doc/giai ma duoc file am thanh ' +
      '(dinh dang khong ho tro, file hong, hoac khong phai WAV/MP3/MP4/M4A hop le)';
    Exit;
  end;
  if SampleCount = 0 then
  begin
    AErrorMsg := 'File am thanh khong co du lieu (0 sample sau khi resample)';
    Exit;
  end;

  // --- Buoc 2: cap phat buffer that va lay du lieu PCM da resample ve
  // 16kHz mono ---
  SetLength(PcmBuffer, SampleCount);
  if not ResampleToWhisperFormat(FileNameW, @PcmBuffer[0], SampleCount, SampleCount) then
  begin
    AErrorMsg := 'WavParser.dll: resample that bai o lan doc du lieu that su ' +
      '(lan hoi kich thuoc truoc do thanh cong - kiem tra lai file co bi doi/xoa giua chung khong)';
    Exit;
  end;

  // --- Buoc 3: transcribe qua transcribe.dll (dung ban Ex de ho tro
  // initial_prompt, tuong thich nguoc voi truong hop AInitialPrompt = '') ---
  ModelPathUtf8 := UTF8String(AModelPath);
  LanguageUtf8 := UTF8String(ALanguage);

  // Uoc luong buffer ket qua: van ban luon nho hon nhieu so voi so sample
  // am thanh goc (16000 sample/giay nhung van ban tieng noi chi vai chuc
  // byte/giay), lay toi thieu 64KB de an toan cho ca clip rat ngan. Dung
  // 1 lan cap phat du lon thay vi goi TranscribeWavFileEx 2 lan (goi 2 lan
  // se chay lai toan bo whisper tu dau, rat lang phi thoi gian).
  OutBufLen := SampleCount;
  if OutBufLen < 65536 then
    OutBufLen := 65536;
  SetLength(OutBuf, OutBufLen);

  // PAnsiChar cua mot UTF8String rong tra ve nil (dac diem Delphi voi
  // chuoi rong) - transcribe.dll da xu ly dung truong hop initialPrompt =
  // NULL (coi nhu khong co goi y), nen khong can if rieng o day.
  InitialPromptUtf8 := UTF8String(AInitialPrompt);
  PromptParam := PAnsiChar(InitialPromptUtf8);

  RetCode := TranscribeWavFileEx(PAnsiChar(ModelPathUtf8), @PcmBuffer[0],
    SampleCount, PAnsiChar(LanguageUtf8), PromptParam, @OutBuf[0], OutBufLen);

  if RetCode < 0 then
  begin
    AErrorMsg := MoTaMaLoiTranscribe(RetCode);
    Exit;
  end;

  // RetCode >= 0: so byte UTF-8 thuc su da ghi vao OutBuf (khong tinh
  // byte null-terminator)
  SetString(RawResult, PAnsiChar(@OutBuf[0]), RetCode);
  AResultText := UTF8ToString(RawResult);
  Result := True;
end;

end.
