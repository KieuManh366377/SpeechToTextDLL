# SpeechToTextDLL

Two lightweight, offline Windows DLLs that add speech-to-text to any VBA, Delphi,
or Win32-compatible application: **WavParser.dll** (decodes WAV/MP3/MP4/M4A to
16kHz mono PCM) and **transcribe.dll** (wraps [whisper.cpp](https://github.com/ggml-org/whisper.cpp)
to turn that PCM into text). Built and tested primarily for Vietnamese, but works
with any language whisper.cpp supports. Runs 100% locally — no audio ever leaves
your machine.

*(Tài liệu chi tiết bằng tiếng Việt ở phần dưới.)*

---

## Giới thiệu

Đây là 2 file DLL độc lập, dùng để thêm tính năng **chuyển giọng nói thành văn bản**
vào bất kỳ ứng dụng Windows nào có thể gọi DLL (VBA/Excel, Delphi/C++ Builder, hoặc
ngôn ngữ khác hỗ trợ gọi hàm C ABI). Chạy hoàn toàn **offline** — audio không bao
giờ được gửi ra ngoài máy tính của bạn.

- **`WavParser.dll`** — đọc và giải mã 4 định dạng audio (**WAV, MP3, MP4, M4A**)
  về đúng chuẩn PCM 16-bit, mono, 16kHz mà mô hình nhận dạng giọng nói cần.
  Tự nhận diện định dạng qua magic-byte đầu file, không dựa vào đuôi file.
- **`transcribe.dll`** — bọc [whisper.cpp](https://github.com/ggml-org/whisper.cpp)
  (viết bằng Go, build qua cgo) để chuyển PCM thành văn bản.

Cả 2 DLL đều build cho **Windows x64**.

## Tính năng

- Hỗ trợ 4 định dạng audio đầu vào: WAV, MP3, MP4, M4A
- Chạy offline hoàn toàn — không có kết nối mạng, không telemetry
- Hỗ trợ nhiều ngôn ngữ (theo model whisper bạn chọn tải về), test kỹ với tiếng Việt
- Có cả đường gọi đồng bộ và bất đồng bộ (polling tiến độ) để không treo UI với file dài
- `initial_prompt` — gợi ý trước từ/tên riêng hay gặp để tăng độ chính xác

## Cài đặt

### Bước 1 — Tải 2 DLL

Tải bản build sẵn từ mục **[Releases](../../releases)** của repo này — không cần
tự build từ source trừ khi bạn muốn tùy chỉnh.

### Bước 2 — Tải model nhận dạng giọng nói (không đi kèm trong repo)

File model (định dạng GGML) **không** được đính kèm trong repo/Releases vì kích
thước quá lớn (vài trăm MB đến vài GB). Kho gốc trên Hugging Face có **rất nhiều**
file model khác nhau (`tiny`, `base`, `small`, `medium`, `large`, bản `.en`, bản
lượng tử hóa `q5`/`q8`...) — dễ tải nhầm nếu chưa quen. Để đơn giản, dùng thẳng 1
trong 3 link tải trực tiếp bên dưới, chọn theo cấu hình máy:

| Model | Phù hợp với | Kích thước | RAM ước tính | Link tải |
|---|---|---|---|---|
| `ggml-small.bin` | **Người dùng phổ thông** — máy cấu hình thường, ưu tiên chạy nhẹ, chấp nhận độ chính xác ở mức khá | ~466 MB | ~852 MB | [Tải trực tiếp](https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-small.bin) |
| `ggml-medium.bin` | Cân bằng — máy tầm trung trở lên, muốn chính xác hơn hẳn small | ~1.53 GB | ~2.1 GB | [Tải trực tiếp](https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-medium.bin) |
| `ggml-large-v3.bin` | **Người dùng chuyên nghiệp** — máy cấu hình mạnh, RAM rộng rãi, ưu tiên độ chính xác cao nhất | ~3.1 GB | ~3.9 GB | [Tải trực tiếp](https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-large-v3.bin) |

Không có "model đúng cho mọi máy" — chọn dựa theo RAM và tốc độ xử lý máy bạn
chấp nhận được. `small` là mức sàn để chương trình chạy được; máy càng khỏe,
càng nên thử `medium`/`large-v3` để có độ chính xác tốt hơn đáng kể. Cần model
khác ngoài 3 lựa chọn trên (ví dụ `tiny` cho máy rất yếu) — xem đầy đủ tại
[kho gốc Hugging Face](https://huggingface.co/ggerganov/whisper.cpp/tree/main).

### Bước 3 — Đặt file

Đặt **cả 2 DLL và model bạn vừa tải** vào **cùng 1 thư mục** với ứng dụng sẽ gọi
chúng (file `.xlsm`, file `.exe`...). Windows tự tìm DLL trong thư mục chứa file
đang chạy, không cần cấu hình PATH.

## Kiểm tra an toàn (VirusTotal)

Cả 2 DLL đã được quét qua [VirusTotal](https://www.virustotal.com) trước khi phát hành:

| File | Kết quả | SHA-256 |
|---|---|---|
| `transcribe.dll` | 0 / 70 | `f9f7985b9f1d7773ce1b11a9d7edc9788980347f3eaec2c13fe1b4edf707fdd5` |
| `WavParser.dll` | 1 / 70 (cảnh báo đơn lẻ, không xác nhận từ các hãng lớn) | `f4f36be6f0e890389ebd8607ff6620078b792d6d94f1c9f469a195d5c2eb6b37` |

Bạn nên tự đối chiếu SHA-256 của file mình tải về với bảng trên (dùng
`certutil -hashfile <file> SHA256` trên Windows) để xác nhận file không bị
chỉnh sửa sau khi phát hành. DLL build từ Go runtime đôi khi bị một số phần
mềm diệt virus báo nhầm (false positive) — nếu antivirus của bạn cảnh báo,
hãy tự kiểm tra lại trên VirusTotal trước khi kết luận.

## Cách sử dụng

### Từ Delphi / C++ Builder

Đây là cách dùng đã được build và test thực tế (xem EXE demo trong repo).

```pascal
uses
  System.SysUtils;

function ResampleToWhisperFormat(lpFileName: PWideChar; pOutBuffer: PSmallInt;
  outBufferCapacity: Cardinal; var pOutSampleCount: Cardinal): LongBool; stdcall;
  external 'WavParser.dll';

function TranscribeWavFile(modelPath: PAnsiChar; pcmSamples: PSmallInt;
  sampleCount: Integer; languageCode: PAnsiChar; outBuf: PAnsiChar;
  outBufLen: Integer): Integer; cdecl; external 'transcribe.dll';

// 1) Hoi truoc so luong sample (pOutBuffer = nil)
var SampleCount: Cardinal;
ResampleToWhisperFormat(PWideChar(FileName), nil, 0, SampleCount);

// 2) Doc du lieu PCM that su
var PcmBuffer: TArray<SmallInt>;
SetLength(PcmBuffer, SampleCount);
ResampleToWhisperFormat(PWideChar(FileName), @PcmBuffer[0], SampleCount, SampleCount);

// 3) Transcribe
var OutBuf: TArray<AnsiChar>;
SetLength(OutBuf, 65536);
var RetCode := TranscribeWavFile(PAnsiChar(UTF8String(ModelPath)), @PcmBuffer[0],
  SampleCount, PAnsiChar(UTF8String('vi')), @OutBuf[0], 65536);
```

Ví dụ đầy đủ hơn (kèm xử lý lỗi, buffer động) xem file `UnitWhisperAPI.pas` trong
thư mục `examples/` của repo.

### Từ VBA (Excel)

**Cách nhanh nhất để thử ngay:** mở sẵn file [`Speech_To_Text_Demo.xlsb`](Speech_To_Text_Demo.xlsb)
đính kèm trong repo — đã nhúng sẵn cả 3 module VBA bên dưới, chỉ cần đặt cùng
thư mục với 2 DLL + model rồi mở lên dùng luôn, không cần tự import gì cả.

Muốn tích hợp vào workbook của riêng bạn: import 3 module VBA dưới đây (`Alt+F11`
→ *File → Import File...*) — đã build và test thực tế, không cần tự viết
`Declare` từ đầu:

| Module | Vai trò |
|---|---|
| [`modDllHelper.bas`](vba/modDllHelper.bas) | Tự load/unload 2 DLL khi mở/đóng workbook (`Auto_Open`/`Auto_Close`), tính sẵn đường dẫn model |
| [`modTestTranscribe.bas`](vba/modTestTranscribe.bas) | Đường **đồng bộ** — `DoTranscribe(WavPath, Language)`, phù hợp file ngắn |
| [`modTranscribeAsync.bas`](vba/modTranscribeAsync.bas) | Đường **bất đồng bộ** — `DoTranscribeAsync(...)`, không treo Excel với file dài, có báo % tiến độ trên StatusBar; kèm sẵn hàm gọi trực tiếp trên ô tính |

Sau khi import xong, dùng ngay trong công thức:

```
=TRANSCRIBE(A1, "vi")
```

với `A1` là đường dẫn file audio (WAV/MP3/MP4/M4A). Hoặc chạy macro
`TestTranscribeToSheet` / `TestTranscribeAsyncToSheet` để chọn file qua hộp
thoại rồi ghi kết quả vào ô đang chọn.

⚠️ **2 điểm dễ vướng khi tự viết code gọi 2 DLL này (đã xử lý sẵn trong 3
module trên, nêu ra để bạn hiểu vì sao code viết như vậy):**
- `WavParser.dll` nhận tên file dạng **Unicode (UTF-16)** — phải dùng `StrPtr()`
  để lấy đúng con trỏ, không truyền thẳng `ByVal ... As String` (VBA sẽ tự ép
  sang ANSI, sai định dạng DLL cần).
- `transcribe.dll` (Go) trả văn bản kết quả dạng **UTF-8** — không dùng
  `StrConv(..., vbUnicode)` để giải mã (hiểu nhầm theo ANSI code page, ra chữ
  rác với tiếng Việt có dấu), phải giải mã qua `MultiByteToWideChar`.

## Tham khảo hàm export

<details>
<summary>WavParser.dll</summary>

| Hàm | Mục đích |
|---|---|
| `GetWavInfo` | Đọc thông tin header WAV (sample rate, channels, bits, data size) |
| `ReadWavPcmData` | Đọc dữ liệu PCM thô từ file WAV |
| `ValidateWavForWhisper` | Kiểm tra nhanh file đã đúng định dạng whisper cần chưa |
| `ResampleToWhisperFormat` | Chuyển bất kỳ định dạng nào (WAV/MP3/MP4/M4A) về PCM 16kHz mono |
| `WriteWavFile` | Ghi buffer PCM ra file .wav hoàn chỉnh (để nghe thử) |

</details>

<details>
<summary>transcribe.dll</summary>

| Hàm | Mục đích |
|---|---|
| `TranscribeWavFile` | Transcribe đồng bộ |
| `TranscribeWavFileWithTimestamps` | Giống trên, có mốc thời gian mỗi câu |
| `TranscribeWavFileEx` | Giống TranscribeWavFile, thêm `initialPrompt` |
| `StartTranscribeAsync` / `...WithTimestamps` | Bắt đầu job nền, không block |
| `GetTranscribeProgress` / `IsTranscribeDone` / `GetTranscribeResult` | Polling tiến độ & lấy kết quả job async |
| `ReleaseModel` | Giải phóng model khỏi cache |

**Mã lỗi:** `-1` không load được model · `-2` lỗi tạo context · `-3` lỗi transcribe ·
số âm khác (không phải các mã trên) buffer quá nhỏ (trị tuyệt đối − 1 = số byte cần) ·
`-6` job async chưa xong · `-10`/`-11` đang có transcribe khác chạy · `-12` mã ngôn
ngữ không hợp lệ · `-13`/`-20`/`-21` tham số NULL/không hợp lệ · `-99` lỗi nội bộ
không lường trước (đã chặn panic, không crash ứng dụng gọi).

</details>

## Giới hạn đã biết

- **Đường bất đồng bộ (`StartTranscribeAsync`)**: nếu buffer bạn truyền vào
  `GetTranscribeResult` quá nhỏ, kết quả job đó sẽ bị mất (không lấy lại được ở
  lần gọi sau) — cấp buffer đủ lớn ngay từ đầu để né vấn đề này. Sẽ vá trong bản
  sau.
- Độ chính xác tiếng Việt phụ thuộc nhiều vào kích thước model — `small` có thể
  nhầm lẫn ở vài phụ âm (ví dụ "gi" nghe ra gần "d"); model lớn hơn cải thiện rõ
  rệt.
- Chưa có sẵn bước lọc nhiễu (denoise/VAD) trước khi transcribe — audio có tạp âm
  nặng sẽ ảnh hưởng độ chính xác.

## Giấy phép & ghi công

Code trong repo này phát hành theo giấy phép **MIT** (xem file `LICENSE`).

DLL sử dụng lại các thư viện mã nguồn mở:
- [whisper.cpp](https://github.com/ggml-org/whisper.cpp) — MIT License
- [minimp3](https://github.com/lieff/minimp3) — public domain / Unlicense
- Windows Media Foundation — API hệ thống Windows, không đóng gói lại

## Tuyên bố miễn trừ trách nhiệm

Phần mềm chia sẻ **miễn phí, "as-is"**, không đi kèm bảo hành dưới bất kỳ hình
thức nào. Tự chịu trách nhiệm khi sử dụng trong môi trường sản xuất/dữ liệu quan
trọng — nên tự kiểm thử kỹ trước khi tích hợp vào quy trình làm việc thật.

## Góp ý / báo lỗi

Repo này chủ yếu để **thăm dò phản hồi thực tế** trước khi phát triển tiếp — rất
hoan nghênh issue báo lỗi, góp ý về độ chính xác trên cấu hình máy/giọng nói khác
nhau, hoặc đề xuất tính năng.
