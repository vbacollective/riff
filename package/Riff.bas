Attribute VB_Name = "Riff"
'/**
' * Riff - Audio Engine (Studio DSP Edition)
' * @description A high-performance, COM-based WASAPI audio engine for VBA (x86/x64 compatible).
' * Contains advanced Array Chunking for zero-latency mixing, Polyphony, and a full
' * Studio DSP Pipeline featuring Freeverb-style Reverb, Chorus, Flanger, Compressor, Biquad EQ, Bitcrusher,
' * RingMod, AutoPan, Delay, BLEP Oscillators, In-Memory Loading, WAV Export, optimized decode v-table calls, Buses, and Peak Meters.
' * @author UesleiDev
' * @version 1.0.7
' */

Option Explicit
Option Private Module

'/** @description Total static audio buffer slots exposed by the engine. */
Private Const RIFF_BUFFER_COUNT As Long = 64

'/** @description Highest valid static audio buffer handle. */
Private Const RIFF_MAX_BUFFER_INDEX As Long = RIFF_BUFFER_COUNT - 1

'/** @description Total polyphonic voice slots exposed by the engine. */
Private Const RIFF_VOICE_COUNT As Long = 32

'/** @description Highest valid polyphonic voice handle. */
Private Const RIFF_MAX_VOICE_INDEX As Long = RIFF_VOICE_COUNT - 1

'/** @description Total audio bus slots exposed by the engine. */
Private Const RIFF_BUS_COUNT As Long = 16

'/** @description Highest valid audio bus index. */
Private Const RIFF_MAX_BUS_INDEX As Long = RIFF_BUS_COUNT - 1



'/**
' * @struct RiffBus
' * @brief Mixer bus state used for grouped volume, fades, mute, solo, and peak metering.
' */
Private Type RiffBus
    Volume As Single
    targetVolume As Single
    FadeStartVolume As Single
    FadeFramesCurrent As Long
    FadeFramesTotal As Long
    PeakL As Single
    PeakR As Single
    Muted As Boolean
    Solo As Boolean
End Type

#If VBA7 Then
    '/** @description Allocates physical memory pages. */
    Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
    
    '/** @description Frees allocated physical memory. */
    Private Declare PtrSafe Function VirtualFree Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal dwFreeType As Long) As Long
    
    '/** @description Fast memory block copy. */
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
    
    '/** @description Memory copy overloaded for Single precision floats. */
    Private Declare PtrSafe Sub RtlMoveMemoryToSingle Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Single, ByVal Source As LongPtr, ByVal Length As LongPtr)
    
    '/** @description Memory copy overloaded for Integers. */
    Private Declare PtrSafe Sub RtlMoveMemoryToInteger Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Integer, ByVal Source As LongPtr, ByVal Length As LongPtr)
    
    '/** @description Clears a block of memory to zeros. */
    Private Declare PtrSafe Sub RtlZeroMemory Lib "kernel32" (ByVal Destination As LongPtr, ByVal Length As LongPtr)
    
    '/** @description Calls a window procedure (used for callbacks). */
    Private Declare PtrSafe Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    
    '/** @description Generic native pointer call shim for optimized x64 COM v-table calls. */
    Private Declare PtrSafe Function RiffCallPtr4 Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As LongPtr, ByVal a0 As LongPtr, ByVal a1 As LongPtr, ByVal a2 As LongPtr, ByVal a3 As LongPtr) As LongPtr
    
    '/** @description Creates a COM object instance. */
    Private Declare PtrSafe Function CoCreateInstance Lib "ole32" (ByRef rclsid As Any, ByVal pUnkOuter As LongPtr, ByVal dwClsContext As Long, ByRef riid As Any, ByRef ppv As LongPtr) As Long
    
    '/** @description Initializes Media Foundation. */
    Private Declare PtrSafe Function MFStartup Lib "mfplat" (ByVal Version As Long, ByVal dwFlags As Long) As Long
    
    '/** @description Shuts down Media Foundation. */
    Private Declare PtrSafe Function MFShutdown Lib "mfplat" () As Long
    
    '/** @description Starts a multimedia timer for the engine loop. */
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    
    '/** @description Kills the multimedia timer. */
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    
    '/** @description Frees memory allocated by COM. */
    Private Declare PtrSafe Function CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr) As Long
    
    '/** @description Low-level COM interface calling via VTable. */
    Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Any, ByRef prgpvarg As Any, ByRef pvargResult As Variant) As Long
    
    '/** @description Converts a String GUID to a structural GUID. */
    Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As Any) As Long
    
    '/** @description Gets a module handle. */
    Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As LongPtr
    
    '/** @description Gets a procedure address from a module. */
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    
    '/** @description Gets a procedure address by ordinal. */
    Private Declare PtrSafe Function GetProcAddressOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As LongPtr, ByVal lpProcName As Long) As LongPtr
    
    '/** @description Creates an MF Source Reader from a file URL. */
    Private Declare PtrSafe Function MFCreateSourceReaderFromURL Lib "mfreadwrite" (ByVal pwszURL As LongPtr, ByVal pAttributes As LongPtr, ByRef ppSourceReader As LongPtr) As Long
    
    '/** @description Creates an empty MF Media Type. */
    Private Declare PtrSafe Function MFCreateMediaType Lib "mfplat" (ByRef ppMFType As LongPtr) As Long
    
    '/** @description Initializes MF Media Type from Wave Format. */
    Private Declare PtrSafe Function MFInitMediaTypeFromWaveFormatEx Lib "mfplat" (ByVal pMFType As LongPtr, ByVal pWaveFormat As LongPtr, ByVal cbBufSize As Long) As Long
    
    '/** @description Creates MF Attributes configuration store. */
    Private Declare PtrSafe Function MFCreateAttributes Lib "mfplat" (ByRef ppMFAttributes As LongPtr, ByVal cInitialSize As Long) As Long
    
    '/** @description Creates an IStream from a memory block. */
    Private Declare PtrSafe Function SHCreateMemStream Lib "shlwapi.dll" (ByVal pInit As LongPtr, ByVal cbInit As Long) As LongPtr
    
    '/** @description Wraps an IStream into an IMFByteStream. */
    Private Declare PtrSafe Function MFCreateMFByteStreamOnStream Lib "mfplat.dll" (ByVal pStream As LongPtr, ByRef ppByteStream As LongPtr) As Long
    
    '/** @description Creates an MF Source Reader directly from an IMFByteStream. */
    Private Declare PtrSafe Function MFCreateSourceReaderFromByteStream Lib "mfreadwrite.dll" (ByVal pByteStream As LongPtr, ByVal pAttributes As LongPtr, ByRef ppSourceReader As LongPtr) As Long
    
    '/** @description Increases system timer resolution. */
    Private Declare PtrSafe Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
    
    '/** @description Restores system timer resolution. */
    Private Declare PtrSafe Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

    '/** @description Returns the address of a Byte() SAFEARRAY pointer slot without indexing the array. */
    Private Declare PtrSafe Function VarPtrByteArray Lib "VBE7.DLL" Alias "VarPtr" (ByRef Var() As Byte) As LongPtr
    
    '/**
    ' * @struct RiffBuffer
    ' * @brief Holds physical memory pointers for static loaded PCM audio data.
    ' */
    Private Type RiffBuffer
        Active As Boolean
        BufferPtr As LongPtr
        BufferLen As Long
    End Type

    '/**
    ' * @struct RiffContext
    ' * @brief Global engine state, holding WASAPI and Media Foundation pointers.
    ' */
    Private Type RiffContext
        Initialized As Boolean
        MagicCookie As Long
        MasterVolume As Single
        MasterPeakL As Single
        MasterPeakR As Single
        sampleRate As Long
        AvgBytesPerSec As Long
        DeviceEnumerator As LongPtr
        Device As LongPtr
        AudioClient As LongPtr
        RenderClient As LongPtr
        MixFormatPtr As LongPtr
        BufferSize As Long
        ThunkTimerCB As LongPtr
        TimerID As LongPtr
        RenderPeriodMs As Long
        MaxWriteFrames As Long
        TimerResolutionActive As Boolean
        TimerCallbackActive As Boolean
        IdleTimerTicks As Long
        AutoSuspendTimer As Boolean
        HasSoloBus As Boolean
        Buffers(0 To RIFF_MAX_BUFFER_INDEX) As RiffBuffer
        Buses(0 To RIFF_MAX_BUS_INDEX) As RiffBus
    End Type
#Else
    '/** @description Allocates physical memory pages. */
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    
    '/** @description Frees allocated physical memory. */
    Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
    
    '/** @description Fast memory block copy. */
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
    
    '/** @description Memory copy overloaded for Single precision floats. */
    Private Declare Sub RtlMoveMemoryToSingle Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Single, ByVal Source As Long, ByVal Length As Long)
    
    '/** @description Memory copy overloaded for Integers. */
    Private Declare Sub RtlMoveMemoryToInteger Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Integer, ByVal Source As Long, ByVal Length As Long)
    
    '/** @description Clears a block of memory to zeros. */
    Private Declare Sub RtlZeroMemory Lib "kernel32" (ByVal Destination As Long, ByVal Length As Long)
    
    '/** @description Calls a window procedure (used for callbacks). */
    Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    '/** @description Generic native pointer call shim for optimized COM v-table calls. */
    Private Declare Function RiffCallPtr4 Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal a0 As Long, ByVal a1 As Long, ByVal a2 As Long, ByVal a3 As Long) As Long
    
    '/** @description Creates a COM object instance. */
    Private Declare Function CoCreateInstance Lib "ole32" (ByRef rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByRef riid As Any, ByRef ppv As Long) As Long
    
    '/** @description Initializes Media Foundation. */
    Private Declare Function MFStartup Lib "mfplat" (ByVal Version As Long, ByVal dwFlags As Long) As Long
    
    '/** @description Shuts down Media Foundation. */
    Private Declare Function MFShutdown Lib "mfplat" () As Long
    
    '/** @description Starts a multimedia timer for the engine loop. */
    Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    
    '/** @description Kills the multimedia timer. */
    Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
    
    '/** @description Frees memory allocated by COM. */
    Private Declare Function CoTaskMemFree Lib "ole32" (ByVal pv As Long) As Long
    
    '/** @description Low-level COM interface calling via VTable. */
    Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Any, ByRef prgpvarg As Any, ByRef pvargResult As Variant) As Long
    
    '/** @description Converts a String GUID to a structural GUID. */
    Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef lpiid As Any) As Long
    
    '/** @description Gets a module handle. */
    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
    
    '/** @description Gets a procedure address from a module. */
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    
    '/** @description Gets a procedure address by ordinal. */
    Private Declare Function GetProcAddressOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As Long) As Long
    
    '/** @description Creates an MF Source Reader from a file URL. */
    Private Declare Function MFCreateSourceReaderFromURL Lib "mfreadwrite" (ByVal pwszURL As Long, ByVal pAttributes As Long, ByRef ppSourceReader As Long) As Long
    
    '/** @description Creates an empty MF Media Type. */
    Private Declare Function MFCreateMediaType Lib "mfplat" (ByRef ppMFType As Long) As Long
    
    '/** @description Initializes MF Media Type from Wave Format. */
    Private Declare Function MFInitMediaTypeFromWaveFormatEx Lib "mfplat" (ByVal pMFType As Long, ByVal pWaveFormat As Long, ByVal cbBufSize As Long) As Long
    
    '/** @description Creates MF Attributes configuration store. */
    Private Declare Function MFCreateAttributes Lib "mfplat" (ByRef ppMFAttributes As Long, ByVal cInitialSize As Long) As Long
    
    '/** @description Creates an IStream from a memory block. */
    Private Declare Function SHCreateMemStream Lib "shlwapi.dll" (ByVal pInit As Long, ByVal cbInit As Long) As Long
    
    '/** @description Wraps an IStream into an IMFByteStream. */
    Private Declare Function MFCreateMFByteStreamOnStream Lib "mfplat.dll" (ByVal pStream As Long, ByRef ppByteStream As Long) As Long
    
    '/** @description Creates an MF Source Reader directly from an IMFByteStream. */
    Private Declare Function MFCreateSourceReaderFromByteStream Lib "mfreadwrite.dll" (ByVal pByteStream As Long, ByVal pAttributes As Long, ByRef ppSourceReader As Long) As Long
    
    '/** @description Increases system timer resolution. */
    Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
    
    '/** @description Restores system timer resolution. */
    Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

    '/** @description Returns the address of a Byte() SAFEARRAY pointer slot without indexing the array. */
    Private Declare Function VarPtrByteArray Lib "VBE6.DLL" Alias "VarPtr" (ByRef Var() As Byte) As Long
    
    '/**
    ' * @struct RiffBuffer
    ' * @brief Holds physical memory pointers for static loaded PCM audio data.
    ' */
    Private Type RiffBuffer
        Active As Boolean
        BufferPtr As Long
        BufferLen As Long
    End Type

    '/**
    ' * @struct RiffContext
    ' * @brief Global engine state, holding WASAPI and Media Foundation pointers.
    ' */
    Private Type RiffContext
        Initialized As Boolean
        MagicCookie As Long
        MasterVolume As Single
        MasterPeakL As Single
        MasterPeakR As Single
        sampleRate As Long
        AvgBytesPerSec As Long
        DeviceEnumerator As Long
        Device As Long
        AudioClient As Long
        RenderClient As Long
        MixFormatPtr As Long
        BufferSize As Long
        ThunkTimerCB As Long
        TimerID As Long
        RenderPeriodMs As Long
        MaxWriteFrames As Long
        TimerResolutionActive As Boolean
        TimerCallbackActive As Boolean
        IdleTimerTicks As Long
        AutoSuspendTimer As Boolean
        HasSoloBus As Boolean
        Buffers(0 To RIFF_MAX_BUFFER_INDEX) As RiffBuffer
        Buses(0 To RIFF_MAX_BUS_INDEX) As RiffBus
    End Type
#End If

'/**
' * @enum RiffWaveType
' * @brief Named oscillator waveform values for playback and offline rendering.
' */
Public Enum RiffWaveType
    RiffWaveSine = 0
    RiffWaveSquare = 1
    RiffWaveSawtooth = 2
    RiffWaveNoise = 3
End Enum

'/**
' * @enum RiffBusId
' * @brief Named audio bus slots for voice routing and bus volume control.
' */
Public Enum RiffBusId
    RiffBusMain = 0
    RiffBusSfx = 1
    RiffBusMusic = 2
    RiffBusVoice = 3
    RiffBusUi = 4
    RiffBusAux1 = 5
    RiffBusAux2 = 6
    RiffBusAux3 = 7
    RiffBusAux4 = 8
    RiffBusAux5 = 9
    RiffBusAux6 = 10
    RiffBusAux7 = 11
    RiffBusAux8 = 12
    RiffBusAux9 = 13
    RiffBusAux10 = 14
    RiffBusAux11 = 15
End Enum

'/**
' * @enum RiffErrorCode
' * @brief Last-error values reported by RiffLastError.
' */
Public Enum RiffErrorCode
    RiffErrorNone = 0
    RiffErrorNotInitialized = 1
    RiffErrorNoFreeBuffer = 2
    RiffErrorNoFreeVoice = 3
    RiffErrorInvalidBuffer = 4
    RiffErrorInvalidVoice = 5
    RiffErrorInvalidBus = 6
    RiffErrorInvalidArgument = 7
    RiffErrorFileNotFound = 8
    RiffErrorComFailure = 9
    RiffErrorMemoryAllocation = 10
    RiffErrorDecodeFailed = 11
    RiffErrorUnsupportedFormat = 12
End Enum

'/**
' * @struct RiffVoice
' * @brief Polyphonic playback channel containing a full Studio DSP Pipeline matrix.
' */
Private Type RiffVoice
    Active As Boolean
    Playing As Boolean
    Paused As Boolean
    busID As RiffBusId
    BufferIndex As Long

    IsOscillator As Boolean
    OscType As RiffWaveType
    OscFreq As Single
    OscPhase As Double
    
    PeakL As Single
    PeakR As Single
    
    Position As Double
    Volume As Single
    Pitch As Double
    Pan As Single
    
    AutoPanRate As Single
    AutoPanDepth As Single
    AutoPanPhase As Double
    
    BitcrushSteps As Single
    BitcrushDownsample As Long
    BitcrushDsCount As Long
    BitcrushLastL As Single
    BitcrushLastR As Single
    
    Distortion As Single
    
    EqBass As Single
    EqMid As Single
    EqTreble As Single
    EqStateLowL As Single
    EqStateLowR As Single
    EqStateHighL As Single
    EqStateHighR As Single
    
    lowPass As Single
    highPass As Single
    FilterStateL As Single
    FilterStateR As Single
    FilterStateHP_L As Single
    FilterStateHP_R As Single
    
    RingModFreq As Single
    RingModMix As Single
    RingModPhase As Double
    
    StereoWidth As Single
    
    TremoloRate As Single
    TremoloDepth As Single
    TremoloPhase As Double
    
    ChorusRate As Single
    ChorusDepth As Single
    ChorusPhase As Double
    
    FlangerRate As Single
    FlangerDepth As Single
    FlangerFeedback As Single
    FlangerPhase As Double
    
    ReverbMix As Single
    ReverbTime As Single
    RevTap1 As Long
    RevTap2 As Long
    RevTap3 As Long
    RevTap4 As Long
    RevDamp1L As Single
    RevDamp1R As Single
    RevDamp2L As Single
    RevDamp2R As Single
    RevDamp3L As Single
    RevDamp3R As Single
    RevDamp4L As Single
    RevDamp4R As Single
    
    BqLowPassZ1L As Single
    BqLowPassZ2L As Single
    BqLowPassZ1R As Single
    BqLowPassZ2R As Single
    BqHighPassZ1L As Single
    BqHighPassZ2L As Single
    BqHighPassZ1R As Single
    BqHighPassZ2R As Single
    EqBassZ1L As Single
    EqBassZ2L As Single
    EqBassZ1R As Single
    EqBassZ2R As Single
    EqMidZ1L As Single
    EqMidZ2L As Single
    EqMidZ1R As Single
    EqMidZ2R As Single
    EqTrebleZ1L As Single
    EqTrebleZ2L As Single
    EqTrebleZ1R As Single
    EqTrebleZ2R As Single
    
    DelayTime As Single
    DelayFeedback As Single
    DelayMix As Single
    
    CompThreshold As Single
    CompRatio As Single
    CompEnv As Single
    
    RingWritePos As Long
    
    Looping As Boolean
    loopStart As Double
    loopEnd As Double
    
    fadeState As Long
    FadeFramesTotal As Long
    FadeFramesCurrent As Long
End Type

'/**
' * @struct GUID
' * @brief Standard Windows COM GUID structure.
' */
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'/** @description Memory allocation flag for VirtualAlloc. */
Private Const MEM_COMMIT As Long = &H1000

'/** @description Memory reservation flag for VirtualAlloc. */
Private Const MEM_RESERVE As Long = &H2000

'/** @description Memory release flag for VirtualFree. */
Private Const MEM_RELEASE As Long = &H8000

'/** @description Read/Write protection flag for Memory. */
Private Const PAGE_READWRITE As Long = &H4

'/** @description Execute/Read/Write protection flag for Thunk Callbacks. */
Private Const PAGE_EXECUTE_READWRITE As Long = &H40

'/** @description Context flag for COM instantiation. */
Private Const CLSCTX_ALL As Long = 23

'/** @description Media Foundation Version constant. */
Private Const MF_VERSION As Long = &H20070

'/** @description Calling convention standard for COM. */
Private Const CC_STDCALL As Long = 4

'/** @description WASAPI Shared Mode constant. */
Private Const AUDCLNT_SHAREMODE_SHARED As Long = 0

'/** @description WASAPI Render endpoint. */
Private Const eRender As Long = 0

'/** @description WASAPI Console role. */
Private Const eConsole As Long = 0

'/** @description Media Foundation stream selection flag. */
Private Const MF_SOURCE_READER_ALL_STREAMS As Long = -2&

'/** @description Media Foundation first audio stream flag. */
Private Const MF_SOURCE_READER_FIRST_AUDIO_STREAM As Long = -3&

'/** @description IUnknown::Release v-table index. */
Private Const VTI_IUNKNOWN_RELEASE As Long = 2

'/** @description IMFAttributes::SetUINT32 v-table index. */
Private Const VTI_MF_ATTRIBUTES_SET_UINT32 As Long = 21

'/** @description IMFSourceReader::SetStreamSelection v-table index. */
Private Const VTI_MF_SOURCE_READER_SET_STREAM_SELECTION As Long = 4

'/** @description IMFSourceReader::SetCurrentMediaType v-table index. */
Private Const VTI_MF_SOURCE_READER_SET_CURRENT_MEDIA_TYPE As Long = 7

'/** @description IMFSourceReader::ReadSample v-table index. */
Private Const VTI_MF_SOURCE_READER_READ_SAMPLE As Long = 9

'/** @description IMFSample::ConvertToContiguousBuffer v-table index. */
Private Const VTI_MF_SAMPLE_CONVERT_TO_CONTIGUOUS_BUFFER As Long = 41

'/** @description IMFMediaBuffer::Lock v-table index. */
Private Const VTI_MF_MEDIA_BUFFER_LOCK As Long = 3

'/** @description IMFMediaBuffer::Unlock v-table index. */
Private Const VTI_MF_MEDIA_BUFFER_UNLOCK As Long = 4

'/** @description IAudioClient::Initialize v-table index. */
Private Const VTI_AUDIO_CLIENT_INITIALIZE As Long = 3

'/** @description IAudioClient::GetBufferSize v-table index. */
Private Const VTI_AUDIO_CLIENT_GET_BUFFER_SIZE As Long = 4

'/** @description IAudioClient::GetCurrentPadding v-table index. */
Private Const VTI_AUDIO_CLIENT_GET_CURRENT_PADDING As Long = 6

'/** @description IAudioClient::IsFormatSupported v-table index. */
Private Const VTI_AUDIO_CLIENT_IS_FORMAT_SUPPORTED As Long = 7

'/** @description IAudioClient::GetMixFormat v-table index. */
Private Const VTI_AUDIO_CLIENT_GET_MIX_FORMAT As Long = 8

'/** @description IAudioClient::Start v-table index. */
Private Const VTI_AUDIO_CLIENT_START As Long = 10

'/** @description IAudioClient::Stop v-table index. */
Private Const VTI_AUDIO_CLIENT_STOP As Long = 11

'/** @description IAudioClient::GetService v-table index. */
Private Const VTI_AUDIO_CLIENT_GET_SERVICE As Long = 14

'/** @description IAudioRenderClient::GetBuffer v-table index. */
Private Const VTI_AUDIO_RENDER_CLIENT_GET_BUFFER As Long = 3

'/** @description IAudioRenderClient::ReleaseBuffer v-table index. */
Private Const VTI_AUDIO_RENDER_CLIENT_RELEASE_BUFFER As Long = 4

'/** @description IMMDevice::Activate v-table index. */
Private Const VTI_MM_DEVICE_ACTIVATE As Long = 3

'/** @description IMMDeviceEnumerator::GetDefaultAudioEndpoint v-table index. */
Private Const VTI_MM_DEVICE_ENUMERATOR_GET_DEFAULT_AUDIO_ENDPOINT As Long = 4

'/** @description Mathematical constant Pi. */
Private Const PI As Double = 3.14159265358979

'/** @description Mathematical constant 2 * Pi. */
Private Const PI2 As Double = 6.28318530717958

'/** @description SAFEARRAY.cDims descriptor offset. */
Private Const RIFF_SAFEARRAY_OFFSET_DIMS As Long = 0

'/** @description SAFEARRAY.cbElements descriptor offset. */
Private Const RIFF_SAFEARRAY_OFFSET_ELEMENT_SIZE As Long = 4

'/** @description 32-bit SAFEARRAY.pvData descriptor offset. */
Private Const RIFF_SAFEARRAY_OFFSET_DATA_32 As Long = 12

'/** @description 64-bit SAFEARRAY.pvData descriptor offset. */
Private Const RIFF_SAFEARRAY_OFFSET_DATA_64 As Long = 16

'/** @description 32-bit SAFEARRAY.rgsabound[0].cElements descriptor offset. */
Private Const RIFF_SAFEARRAY_OFFSET_COUNT_32 As Long = 16

'/** @description 64-bit SAFEARRAY.rgsabound[0].cElements descriptor offset. */
Private Const RIFF_SAFEARRAY_OFFSET_COUNT_64 As Long = 24

'/** @description Low-latency render tick interval in milliseconds. */
Private Const RIFF_RENDER_PERIOD_MS As Long = 20

'/** @description Number of idle timer ticks before the render timer suspends itself to release the VBE. */
Private Const RIFF_IDLE_TIMER_STOP_TICKS As Long = 50

'/** @description Maximum audio chunk written per timer tick in milliseconds. Keeps enough headroom to avoid chopped playback. */
Private Const RIFF_MAX_WRITE_MS As Long = 100

'/** @description Requested WASAPI shared buffer duration in milliseconds. Avoids underruns while still preventing a long silent queue. */
Private Const RIFF_DEVICE_BUFFER_MS As Long = 150

'/** @description WASAPI REFERENCE_TIME units per millisecond. */
Private Const RIFF_HNS_PER_MS As Long = 10000

'/** @description Milliseconds per second. */
Private Const RIFF_MS_PER_SEC As Long = 1000

'/** @description Requested multimedia timer resolution in milliseconds. */
Private Const RIFF_TIMER_RESOLUTION_MS As Long = 1

'/** @description Timer callback guard value, ASCII "RIFF". */
Private Const RIFF_MAGIC_COOKIE As Long = &H52494646

'/** @description Per-voice stereo ring buffer capacity in samples. */
Private Const RIFF_RING_SAMPLES_PER_VOICE As Long = 192000

'/** @description Initial native decode scratch allocation size in bytes. */
Private Const RIFF_DECODE_INITIAL_CAPACITY As Long = 1048576

'/** @description Native decode scratch growth multiplier when capacity is exhausted. */
Private Const RIFF_DECODE_GROWTH_FACTOR As Long = 2

'/** @description Byte widths for strongly typed sample buffers. */
Private Const RIFF_SINGLE_BYTES As Long = 4
Private Const RIFF_LONG_BYTES As Long = 4
Private Const RIFF_INTEGER_BYTES As Long = 2
Private Const RIFF_BYTE_BYTES As Long = 1

'/** @description Number of output channels written by WAV export helpers. */
Private Const RIFF_WAV_EXPORT_CHANNELS As Integer = 2

'/** @description Common normalized sample/control range. */
Private Const RIFF_UNITY_GAIN As Single = 1!
Private Const RIFF_NEGATIVE_UNITY_GAIN As Single = -1!
Private Const RIFF_HALF_SCALE As Single = 0.5!
Private Const RIFF_PHASE_FULL As Double = 1#
Private Const RIFF_PHASE_HALF As Double = 0.5
Private Const RIFF_BIPOLAR_DOUBLE_SCALE As Double = 2#
Private Const RIFF_BIPOLAR_SINGLE_SCALE As Single = 2!

'/** @description Bits per sample written by WAV export helpers. */
Private Const RIFF_WAV_EXPORT_BITS As Integer = 16

'/** @description Supported PCM/native sample bit depths. */
Private Const RIFF_PCM16_BITS As Integer = 16
Private Const RIFF_PCM24_BITS As Integer = 24
Private Const RIFF_FLOAT32_BITS As Integer = 32

'/** @description Bytes per exported stereo PCM16 frame. */
Private Const RIFF_WAV_EXPORT_FRAME_BYTES As Long = 4

'/** @description Standard PCM fmt chunk payload size. */
Private Const RIFF_WAV_FMT_CHUNK_SIZE As Long = 16

'/** @description Standard PCM WAV format tag. */
Private Const RIFF_WAV_FORMAT_PCM As Integer = 1

'/** @description RIFF chunk payload overhead before data bytes. */
Private Const RIFF_WAV_SIZE_OVERHEAD As Long = 36

'/** @description PCM16 positive full-scale value. */
Private Const RIFF_PCM16_MAX As Long = 32767

'/** @description PCM16 negative full-scale magnitude. */
Private Const RIFF_PCM16_MIN_MAGNITUDE As Long = 32768

'/** @description PCM16 signed minimum value. */
Private Const RIFF_PCM16_MIN As Long = -32768

'/** @description PCM16-to-float normalization multiplier. */
Private Const RIFF_PCM16_TO_FLOAT_SCALE As Single = 3.051758E-05!

'/** @description PCM24 signed normalization denominator. */
Private Const RIFF_PCM24_SCALE As Double = 8388608#

'/** @description PCM24 byte shift constants. */
Private Const RIFF_BYTE_SHIFT_8 As Long = &H100&
Private Const RIFF_BYTE_SHIFT_16 As Long = &H10000
Private Const RIFF_PCM24_SIGN_BIT As Byte = &H80
Private Const RIFF_PCM24_SIGN_EXTEND As Long = -16777216

'/** @description PCM32 signed normalization denominator. */
Private Const RIFF_PCM32_SCALE As Double = 2147483648#

'/** @description PCM32 positive full-scale value used before integer output conversion. */
Private Const RIFF_PCM32_MAX As Double = 2147483647#

'/** @description WAVEFORMATEX format tag offset. */
Private Const RIFF_WFX_FORMAT_TAG_OFFSET As Long = 0

'/** @description WAVEFORMATEX channel count offset. */
Private Const RIFF_WFX_CHANNELS_OFFSET As Long = 2

'/** @description WAVEFORMATEX sample-rate offset. */
Private Const RIFF_WFX_SAMPLE_RATE_OFFSET As Long = 4

'/** @description WAVEFORMATEX average-bytes-per-second offset. */
Private Const RIFF_WFX_AVG_BYTES_OFFSET As Long = 8

'/** @description WAVEFORMATEX block-align offset. */
Private Const RIFF_WFX_BLOCK_ALIGN_OFFSET As Long = 12

'/** @description WAVEFORMATEX bits-per-sample offset. */
Private Const RIFF_WFX_BITS_OFFSET As Long = 14

'/** @description WAVEFORMATEX extension-size offset. */
Private Const RIFF_WFX_CB_SIZE_OFFSET As Long = 16

'/** @description WAVEFORMATEXTENSIBLE valid-bits offset. */
Private Const RIFF_WFX_VALID_BITS_OFFSET As Long = 18

'/** @description WAVEFORMATEXTENSIBLE channel-mask offset. */
Private Const RIFF_WFX_CHANNEL_MASK_OFFSET As Long = 20

'/** @description WAVEFORMATEXTENSIBLE subformat GUID Data1 offset. */
Private Const RIFF_WFX_SUBFORMAT_OFFSET As Long = 24

'/** @description WAVEFORMATEXTENSIBLE subformat GUID Data2 offset. */
Private Const RIFF_WFX_SUBFORMAT_DATA2_OFFSET As Long = 28

'/** @description WAVEFORMATEXTENSIBLE subformat GUID Data3 offset. */
Private Const RIFF_WFX_SUBFORMAT_DATA3_OFFSET As Long = 30

'/** @description WAVEFORMATEXTENSIBLE subformat GUID Data4 offset. */
Private Const RIFF_WFX_SUBFORMAT_DATA4_OFFSET As Long = 32

'/** @description Base WAVEFORMATEX byte size. */
Private Const RIFF_WFX_BASE_SIZE As Long = 18

'/** @description WAVEFORMATEXTENSIBLE extension byte size. */
Private Const RIFF_WFX_EXTENSIBLE_CB_SIZE As Integer = 22

'/** @description IEEE float WAVEFORMATEX tag and subformat Data1 value. */
Private Const RIFF_WAVE_FORMAT_IEEE_FLOAT As Long = 3

'/** @description IEEE float WAVEFORMATEXTENSIBLE subformat GUID values. */
Private Const RIFF_IEEE_FLOAT_SUBFORMAT_DATA2 As Integer = 0
Private Const RIFF_IEEE_FLOAT_SUBFORMAT_DATA3 As Integer = 16
Private Const RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_0 As Byte = 128
Private Const RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_1 As Byte = 0
Private Const RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_2 As Byte = 0
Private Const RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_3 As Byte = 170
Private Const RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_4 As Byte = 0
Private Const RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_5 As Byte = 56
Private Const RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_6 As Byte = 155
Private Const RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_7 As Byte = 113
Private Const RIFF_GUID_DATA4_BYTES As Long = 8

'/** @description WAVE_FORMAT_EXTENSIBLE tag as signed Integer. */
Private Const RIFF_WAVE_FORMAT_EXTENSIBLE As Integer = -2

'/** @description Stereo channel mask for front-left/front-right. */
Private Const RIFF_STEREO_CHANNEL_MASK As Long = 3

'/** @description Default sample rate used only when a device reports an invalid value during format promotion. */
Private Const RIFF_DEFAULT_SAMPLE_RATE As Long = 44100

'/** @description Shared-mode initialize flag for event callback mode. */
Private Const AUDCLNT_STREAMFLAGS_EVENTCALLBACK As Long = &H80000000

'/** @description Media Foundation end-of-stream flag returned by IMFSourceReader::ReadSample. */
Private Const MF_SOURCE_READERF_ENDOFSTREAM As Long = 2

'/** @description Initial Media Foundation attributes capacity used by source readers. */
Private Const RIFF_MF_ATTRIBUTES_CAPACITY As Long = 1

'/** @description Media Foundation stream-selection boolean values. */
Private Const RIFF_MF_STREAM_DESELECT As Long = 0
Private Const RIFF_MF_STREAM_SELECT As Long = 1

'/** @description Media Foundation source reader resampler attribute GUID. */
Private Const RIFF_GUID_MF_SOURCE_READER_ENABLE_ADVANCED_VIDEO_PROCESSING As String = "{7632CB14-D379-4770-AE7D-EA24154D9298}"

'/** @description Freeverb tap delay multipliers in seconds. */
Private Const RIFF_REVERB_TAP1_SEC As Single = 0.0297!
Private Const RIFF_REVERB_TAP2_SEC As Single = 0.0371!
Private Const RIFF_REVERB_TAP3_SEC As Single = 0.0411!
Private Const RIFF_REVERB_TAP4_SEC As Single = 0.0437!

'/** @description Freeverb wet mix weights. */
Private Const RIFF_REVERB_DIRECT_WET As Single = 0.19!
Private Const RIFF_REVERB_CROSS_WET As Single = 0.055!

'/** @description Peak meter decay multipliers. */
Private Const RIFF_MASTER_IDLE_PEAK_DECAY As Single = 0.85!
Private Const RIFF_ACTIVE_PEAK_DECAY As Single = 0.9!

'/** @description Safety clamp for feedback-style effects. */
Private Const RIFF_MAX_FEEDBACK As Single = 0.95!

'/** @description Safety clamp for stereo width. */
Private Const RIFF_MAX_STEREO_WIDTH As Single = 0.99!

'/** @description Low-pass bypass threshold. */
Private Const RIFF_LOWPASS_BYPASS_THRESHOLD As Single = 0.999!

'/** @description Default EQ crossover frequencies used by the simple 3-band tone stage. */
Private Const RIFF_EQ_LOW_CROSSOVER_HZ As Single = 200!
Private Const RIFF_EQ_HIGH_CROSSOVER_HZ As Single = 2000!

'/** @description Default parametric EQ center frequencies. */
Private Const RIFF_EQ_BASS_CENTER_HZ As Single = 120!
Private Const RIFF_EQ_MID_CENTER_HZ As Single = 1000!
Private Const RIFF_EQ_TREBLE_CENTER_HZ As Single = 6500!

'/** @description Default parametric EQ Q value. */
Private Const RIFF_EQ_DEFAULT_Q As Single = 0.7!
Private Const RIFF_EQ_MID_Q As Single = 1!
Private Const RIFF_BIQUAD_MIN_Q As Single = 0.1!
Private Const RIFF_BIQUAD_MAX_Q As Single = 12!
Private Const RIFF_BIQUAD_MIN_GAIN As Single = 0.05!
Private Const RIFF_BIQUAD_MAX_GAIN As Single = 8!

'/** @description Default biquad Q used by simple low/high-pass voice filters. */
Private Const RIFF_FILTER_DEFAULT_Q As Single = 0.707!

'/** @description Voice filter cutoff limits and Nyquist ratios. */
Private Const RIFF_FILTER_MIN_CUTOFF_HZ As Single = 20!
Private Const RIFF_LOWPASS_MIN_CUTOFF_HZ As Single = 40!
Private Const RIFF_LOWPASS_MAX_SAMPLE_RATE_RATIO As Single = 0.45!
Private Const RIFF_HIGHPASS_MAX_SAMPLE_RATE_RATIO As Single = 0.35!

'/** @description Modulated delay ranges in seconds. */
Private Const RIFF_FLANGER_BASE_DELAY_SEC As Single = 0.002!
Private Const RIFF_FLANGER_MOD_DELAY_SEC As Single = 0.005!
Private Const RIFF_CHORUS_BASE_DELAY_SEC As Single = 0.02!
Private Const RIFF_CHORUS_MOD_DELAY_SEC As Single = 0.005!

'/** @description Built-in oscillator defaults. */
Private Const RIFF_DEFAULT_OSCILLATOR_HZ As Single = 440!
Private Const RIFF_OSCILLATOR_EXPORT_GAIN As Single = 0.75!
Private Const RIFF_OSCILLATOR_FALLBACK_HZ As Double = 440#
Private Const RIFF_OSCILLATOR_MAX_DT As Double = 0.5
Private Const RIFF_OSCILLATOR_BLEP_LEVEL As Single = 0.65!

'/** @description Default voice effect values. */
Private Const RIFF_DEFAULT_VOICE_VOLUME As Single = 1!
Private Const RIFF_DEFAULT_VOICE_PITCH As Double = 1#
Private Const RIFF_DEFAULT_DISTORTION As Single = 1!
Private Const RIFF_DEFAULT_FILTER_CONTROL As Single = 1!
Private Const RIFF_DEFAULT_BITCRUSH_DOWNSAMPLE As Long = 1
Private Const RIFF_DEFAULT_CHORUS_RATE As Single = 1.5!
Private Const RIFF_DEFAULT_FLANGER_RATE As Single = 0.5!
Private Const RIFF_DEFAULT_REVERB_TIME As Single = 0.5!

'/** @description Public control clamp limits. */
Private Const RIFF_MAX_BUS_VOLUME As Single = 2!
Private Const RIFF_MIN_VOICE_PITCH As Single = 0.1!
Private Const RIFF_MAX_EQ_GAIN As Single = 5!
Private Const RIFF_MIN_MOD_RATE_HZ As Single = 0.1!
Private Const RIFF_MAX_MOD_RATE_HZ As Single = 10!
Private Const RIFF_MAX_STEREO_WIDTH_CONTROL As Single = 5!
Private Const RIFF_MAX_BIT_DEPTH As Single = 32!
Private Const RIFF_MIN_BIT_DEPTH As Single = 2!
Private Const RIFF_MIN_DOWNSAMPLE_FACTOR As Long = 1
Private Const RIFF_MAX_LFO_RATE_HZ As Single = 20!
Private Const RIFF_MAX_COMPRESSOR_RATIO As Single = 20!
Private Const RIFF_MIN_COMPRESSOR_RATIO As Single = 1!
Private Const RIFF_MIN_COMPRESSOR_THRESHOLD As Single = 0.01!
Private Const RIFF_MIN_LOWPASS_CONTROL As Single = 0.01!

'/** @description Compressor envelope smoothing constants. */
Private Const RIFF_COMP_ENV_FLOOR As Single = 0.0001!
Private Const RIFF_COMP_ATTACK_COEFF As Single = 0.01!
Private Const RIFF_COMP_RELEASE_COEFF As Single = 0.001!

'/** @description Freeverb feedback and damping shaping coefficients. */
Private Const RIFF_REVERB_FEEDBACK_BASE As Single = 0.68!
Private Const RIFF_REVERB_FEEDBACK_RANGE As Single = 0.28!
Private Const RIFF_REVERB_DAMP_BASE As Single = 0.18!
Private Const RIFF_REVERB_DAMP_RANGE As Single = 0.24!

'/** @description Hex byte string width used when decoding generated thunk opcodes. */
Private Const RIFF_HEX_BYTE_CHARS As Long = 2

'/** @description x64 thunk placeholder offsets. */
Private Const RIFF_THUNK64_SIZE As Long = 1024
Private Const RIFF_THUNK64_EBMODE_OFFSET As Long = 26
Private Const RIFF_THUNK64_KILLTIMER_OFFSET As Long = 63
Private Const RIFF_THUNK64_CALLBACK_OFFSET As Long = 102

'/** @description x86 thunk placeholder offsets. */
Private Const RIFF_THUNK32_SIZE As Long = 512
Private Const RIFF_THUNK32_EBMODE_OFFSET As Long = 4
Private Const RIFF_THUNK32_KILLTIMER_OFFSET As Long = 33
Private Const RIFF_THUNK32_CALLBACK_OFFSET As Long = 62

'/** @description VBE EbMode export ordinal used by generated timer thunks. */
Private Const RIFF_VBE_EBMODE_ORDINAL As Long = 1

'/** @description WASAPI COM class/interface GUIDs. */
Private Const RIFF_GUID_MM_DEVICE_ENUMERATOR_CLASS As String = "{A95664D2-9614-4F35-A746-DE8DB63617E6}"
Private Const RIFF_GUID_AUDIO_RENDER_CLIENT As String = "{F294ACFC-3146-4483-A7BF-ADDCA7C260E2}"

'/** @description Global state holding hardware info and context. */
Private rCtx As RiffContext

'/** @description Last public API failure reported by RiffLastError. */
Private rLastError As RiffErrorCode

'/** @description Pool of 32 polyphonic voices for audio playback. */
Private rVoices(0 To RIFF_MAX_VOICE_INDEX) As RiffVoice

'/**
' * @description Contiguous global ring buffer array for spatial effects.
' * Solves the Column-Major 2D wipe issue by providing 1D sequential access.
' */
Private rRingBuf() As Single
Private rMixArr32() As Single
Private rMixInt32() As Long
Private rSrcArr32() As Single
Private rSrcArrI32() As Long
Private rMixArr16() As Integer
Private rSrcArr16() As Integer
Private rMixArr32Cap As Long
Private rMixInt32Cap As Long
Private rSrcArr32Cap As Long
Private rSrcArrI32Cap As Long
Private rMixArr16Cap As Long
Private rSrcArr16Cap As Long

'/**
' * @property RiffLastError
' * @brief Returns the most recent RiffErrorCode set by a public API call.
' */
Public Property Get RiffLastError() As RiffErrorCode
    RiffLastError = rLastError
End Property

'/**
' * @function RiffSetLastError
' * @brief Stores the current public API failure state.
' */
Private Sub RiffSetLastError(ByVal errorCode As RiffErrorCode)
    rLastError = errorCode
End Sub

'/**
' * @function RiffRequireInitialized
' * @brief Returns True when the engine is initialized and records an error otherwise.
' */
Private Function RiffRequireInitialized() As Boolean
    If rCtx.Initialized Then
        RiffSetLastError RiffErrorNone
        RiffRequireInitialized = True
    Else
        RiffSetLastError RiffErrorNotInitialized
    End If
End Function

'/**
' * @function RiffIsValidBufferHandle
' * @brief Validates a static buffer handle range.
' */
Private Function RiffIsValidBufferHandle(ByVal bufferHandle As Long) As Boolean
    RiffIsValidBufferHandle = (bufferHandle >= 0 And bufferHandle <= RIFF_MAX_BUFFER_INDEX)
End Function

'/**
' * @function RiffIsValidVoiceHandle
' * @brief Validates a voice handle range.
' */
Private Function RiffIsValidVoiceHandle(ByVal voiceHandle As Long) As Boolean
    RiffIsValidVoiceHandle = (voiceHandle >= 0 And voiceHandle <= RIFF_MAX_VOICE_INDEX)
End Function

'/**
' * @function RiffRequireBufferHandle
' * @brief Validates engine initialization and a buffer handle range.
' */
Private Function RiffRequireBufferHandle(ByVal bufferHandle As Long) As Boolean
    If Not RiffRequireInitialized() Then
        Exit Function
    End If

    If Not RiffIsValidBufferHandle(bufferHandle) Then
        RiffSetLastError RiffErrorInvalidBuffer
        Exit Function
    End If

    RiffRequireBufferHandle = True
End Function

'/**
' * @function RiffRequireVoiceHandle
' * @brief Validates engine initialization and a voice handle range.
' */
Private Function RiffRequireVoiceHandle(ByVal voiceHandle As Long) As Boolean
    If Not RiffRequireInitialized() Then
        Exit Function
    End If

    If Not RiffIsValidVoiceHandle(voiceHandle) Then
        RiffSetLastError RiffErrorInvalidVoice
        Exit Function
    End If

    RiffRequireVoiceHandle = True
End Function

'/**
' * @function RiffClampBusVolume
' * @brief Clamps a bus volume multiplier to the safe public bus range.
' * @param value Input volume multiplier.
' * @return {Single} Clamped volume multiplier.
' */
Private Function RiffClampBusVolume(ByVal value As Single) As Single
    If value < 0! Then
        RiffClampBusVolume = 0!
    ElseIf value > RIFF_MAX_BUS_VOLUME Then
        RiffClampBusVolume = RIFF_MAX_BUS_VOLUME
    Else
        RiffClampBusVolume = value
    End If
End Function

'/**
' * @function RiffResetBusState
' * @brief Restores a bus to neutral mixer state.
' * @param busID The target bus index.
' */
Private Sub RiffResetBusState(ByVal busID As Long)
    If busID < 0 Or busID > RIFF_MAX_BUS_INDEX Then
        Exit Sub
    End If

    rCtx.Buses(busID).Volume = RIFF_UNITY_GAIN
    rCtx.Buses(busID).targetVolume = RIFF_UNITY_GAIN
    rCtx.Buses(busID).FadeStartVolume = RIFF_UNITY_GAIN
    rCtx.Buses(busID).FadeFramesCurrent = 0
    rCtx.Buses(busID).FadeFramesTotal = 0
    rCtx.Buses(busID).PeakL = 0!
    rCtx.Buses(busID).PeakR = 0!
    rCtx.Buses(busID).Muted = False
    rCtx.Buses(busID).Solo = False
End Sub

'/**
' * @function RiffRefreshSoloState
' * @brief Recomputes the fast global flag used to decide whether solo routing is active.
' */
Private Sub RiffRefreshSoloState()
    Dim i As Long

    rCtx.HasSoloBus = False
    For i = 0 To RIFF_MAX_BUS_INDEX
        If rCtx.Buses(i).Solo Then
            rCtx.HasSoloBus = True
            Exit Sub
        End If
    Next i
End Sub

'/**
' * @function RiffTickBusFades
' * @brief Advances all active bus fades by the requested number of rendered frames.
' * @param renderedFrames Number of audio frames being rendered this tick.
' */
Private Sub RiffTickBusFades(ByVal renderedFrames As Long)
    Dim i As Long
    Dim t As Single

    If renderedFrames <= 0 Then
        Exit Sub
    End If

    For i = 0 To RIFF_MAX_BUS_INDEX
        If rCtx.Buses(i).FadeFramesTotal > 0 Then
            rCtx.Buses(i).FadeFramesCurrent = rCtx.Buses(i).FadeFramesCurrent + renderedFrames
            If rCtx.Buses(i).FadeFramesCurrent >= rCtx.Buses(i).FadeFramesTotal Then
                rCtx.Buses(i).Volume = rCtx.Buses(i).targetVolume
                rCtx.Buses(i).FadeFramesCurrent = 0
                rCtx.Buses(i).FadeFramesTotal = 0
                rCtx.Buses(i).FadeStartVolume = rCtx.Buses(i).Volume
            Else
                t = CSng(rCtx.Buses(i).FadeFramesCurrent) / CSng(rCtx.Buses(i).FadeFramesTotal)
                rCtx.Buses(i).Volume = rCtx.Buses(i).FadeStartVolume + ((rCtx.Buses(i).targetVolume - rCtx.Buses(i).FadeStartVolume) * t)
            End If
        End If

        rCtx.Buses(i).PeakL = rCtx.Buses(i).PeakL * RIFF_ACTIVE_PEAK_DECAY
        rCtx.Buses(i).PeakR = rCtx.Buses(i).PeakR * RIFF_ACTIVE_PEAK_DECAY
    Next i
End Sub

'/**
' * @function RiffBusMixVolume
' * @brief Returns the effective bus volume used by the mixer after mute and solo routing.
' * @param busID The target bus index.
' * @return {Single} Effective gain for this render pass.
' */
Private Function RiffBusMixVolume(ByVal busID As RiffBusId) As Single
    If busID < RiffBusMain Or busID > RIFF_MAX_BUS_INDEX Then
        RiffBusMixVolume = 0!
        Exit Function
    End If

    If rCtx.Buses(busID).Muted Then
        RiffBusMixVolume = 0!
    ElseIf rCtx.HasSoloBus And Not rCtx.Buses(busID).Solo Then
        RiffBusMixVolume = 0!
    Else
        RiffBusMixVolume = rCtx.Buses(busID).Volume
    End If
End Function

'/**
' * @function RiffClampBusId
' * @brief Clamps an enum-backed bus value to the valid bus range.
' */
Private Function RiffClampBusId(ByVal busID As RiffBusId) As RiffBusId
    If busID < RiffBusMain Then
        RiffSetLastError RiffErrorInvalidBus
        RiffClampBusId = RiffBusMain
    ElseIf busID > RIFF_MAX_BUS_INDEX Then
        RiffSetLastError RiffErrorInvalidBus
        RiffClampBusId = RIFF_MAX_BUS_INDEX
    Else
        RiffClampBusId = busID
    End If
End Function

'/**
' * @function RiffClampWaveType
' * @brief Clamps an enum-backed waveform value to the valid waveform range.
' */
Private Function RiffClampWaveType(ByVal waveType As RiffWaveType) As RiffWaveType
    If waveType < RiffWaveSine Then
        RiffSetLastError RiffErrorInvalidArgument
        RiffClampWaveType = RiffWaveSine
    ElseIf waveType > RiffWaveNoise Then
        RiffSetLastError RiffErrorInvalidArgument
        RiffClampWaveType = RiffWaveNoise
    Else
        RiffClampWaveType = waveType
    End If
End Function

'/**
' * @function RiffTryGetByteArrayData
' * @brief Reads a VBA Byte() SAFEARRAY descriptor without raising on uninitialized arrays.
' */
#If VBA7 Then
Private Function RiffTryGetByteArrayData(ByRef audioData() As Byte, ByRef dataPtr As LongPtr, ByRef byteCount As Long) As Boolean
    Dim pArraySlot As LongPtr
    Dim pSafeArray As LongPtr
    Dim cDims As Integer
    Dim cbElements As Long
    Dim cElements As Long

    pArraySlot = VarPtrByteArray(audioData)
    If pArraySlot = 0 Then
        Exit Function
    End If

    RtlMoveMemory VarPtr(pSafeArray), ByVal pArraySlot, LenB(pSafeArray)
    If pSafeArray = 0 Then
        Exit Function
    End If

    RtlMoveMemory VarPtr(cDims), ByVal (pSafeArray + RIFF_SAFEARRAY_OFFSET_DIMS), LenB(cDims)
    RtlMoveMemory VarPtr(cbElements), ByVal (pSafeArray + RIFF_SAFEARRAY_OFFSET_ELEMENT_SIZE), LenB(cbElements)
    RtlMoveMemory VarPtr(dataPtr), ByVal (pSafeArray + RIFF_SAFEARRAY_OFFSET_DATA_64), LenB(dataPtr)
    RtlMoveMemory VarPtr(cElements), ByVal (pSafeArray + RIFF_SAFEARRAY_OFFSET_COUNT_64), LenB(cElements)

    If cDims <> 1 Then
        Exit Function
    End If
    If cbElements <> 1 Then
        Exit Function
    End If
    If dataPtr = 0 Or cElements <= 0 Then
        Exit Function
    End If

    byteCount = cElements
    RiffTryGetByteArrayData = True
End Function
#Else
Private Function RiffTryGetByteArrayData(ByRef audioData() As Byte, ByRef dataPtr As Long, ByRef byteCount As Long) As Boolean
    Dim pArraySlot As Long
    Dim pSafeArray As Long
    Dim cDims As Integer
    Dim cbElements As Long
    Dim cElements As Long

    pArraySlot = VarPtrByteArray(audioData)
    If pArraySlot = 0 Then
        Exit Function
    End If

    RtlMoveMemory VarPtr(pSafeArray), ByVal pArraySlot, LenB(pSafeArray)
    If pSafeArray = 0 Then
        Exit Function
    End If

    RtlMoveMemory VarPtr(cDims), ByVal (pSafeArray + RIFF_SAFEARRAY_OFFSET_DIMS), LenB(cDims)
    RtlMoveMemory VarPtr(cbElements), ByVal (pSafeArray + RIFF_SAFEARRAY_OFFSET_ELEMENT_SIZE), LenB(cbElements)
    RtlMoveMemory VarPtr(dataPtr), ByVal (pSafeArray + RIFF_SAFEARRAY_OFFSET_DATA_32), LenB(dataPtr)
    RtlMoveMemory VarPtr(cElements), ByVal (pSafeArray + RIFF_SAFEARRAY_OFFSET_COUNT_32), LenB(cElements)

    If cDims <> 1 Then
        Exit Function
    End If
    If cbElements <> 1 Then
        Exit Function
    End If
    If dataPtr = 0 Or cElements <= 0 Then
        Exit Function
    End If

    byteCount = cElements
    RiffTryGetByteArrayData = True
End Function
#End If


'/**
' * @function RiffOpen
' * @brief Initializes the WASAPI Audio Engine, Media Foundation, and background Timers.
' * @return {Boolean} True if initialization was successful.
' */
Public Function RiffOpen() As Boolean
    RiffSetLastError RiffErrorNone
    If rCtx.Initialized Then
        RiffOpen = True
        Exit Function
    End If

    rCtx.MagicCookie = RIFF_MAGIC_COOKIE
    rCtx.MasterVolume = RIFF_UNITY_GAIN
    rCtx.MasterPeakL = 0!
    rCtx.MasterPeakR = 0!
    rCtx.RenderPeriodMs = RIFF_RENDER_PERIOD_MS
    rCtx.MaxWriteFrames = 0

    Dim i As Long
    For i = 0 To RIFF_MAX_BUS_INDEX
        RiffResetBusState i
    Next i
    rCtx.HasSoloBus = False

    If Not InitThunks() Then
        RiffSetLastError RiffErrorMemoryAllocation
        Exit Function
    End If

    If MFStartup(MF_VERSION, 0) <> 0 Then
        FreeThunks
        RiffSetLastError RiffErrorComFailure
        Exit Function
    End If

    If Not InitWASAPI() Then
        ReleaseWASAPI
        MFShutdown
        FreeThunks
        RiffSetLastError RiffErrorComFailure
        Exit Function
    End If

    ReDim rRingBuf(0 To (RIFF_VOICE_COUNT * RIFF_RING_SAMPLES_PER_VOICE) + 1)

    rCtx.Initialized = True
    RiffOpen = True
End Function

'/**
' * @function RiffClose
' * @brief Safely shuts down the audio engine, frees memory, and releases COM objects.
' */
Public Sub RiffClose()
    If Not rCtx.Initialized Then
        RiffSetLastError RiffErrorNotInitialized
        Exit Sub
    End If
    
    rCtx.MagicCookie = 0
    
    RiffStopRenderTimer
    
    Dim i As Long
    For i = 0 To RIFF_MAX_VOICE_INDEX
        rVoices(i).Active = False
    Next i

    For i = 0 To RIFF_MAX_BUFFER_INDEX
        If rCtx.Buffers(i).Active Then
            RiffUnload i
        End If
    Next i
    
    Erase rRingBuf
    Erase rMixArr32
    Erase rMixInt32
    Erase rSrcArr32
    Erase rSrcArrI32
    Erase rMixArr16
    Erase rSrcArr16
    rMixArr32Cap = 0
    rMixInt32Cap = 0
    rSrcArr32Cap = 0
    rSrcArrI32Cap = 0
    rMixArr16Cap = 0
    rSrcArr16Cap = 0

    ReleaseWASAPI
    MFShutdown
    FreeThunks
    
    rCtx.Initialized = False
    RiffSetLastError RiffErrorNone
End Sub


'/**
' * @property RiffIsInitialized
' * @brief Returns True if the audio engine is running.
' */
Public Property Get RiffIsInitialized() As Boolean
    RiffIsInitialized = rCtx.Initialized
End Property

'/**
' * @property RiffAutoSuspendTimer
' * @brief Gets or sets whether the render timer automatically stops while no voices are active.
' */
Public Property Get RiffAutoSuspendTimer() As Boolean
    RiffAutoSuspendTimer = rCtx.AutoSuspendTimer
End Property

'/**
' * @property RiffAutoSuspendTimer
' * @brief Enables or disables automatic render timer suspension when playback is idle.
' * @param value True to suspend the timer while idle, False to keep it alive until RiffClose.
' */
Public Property Let RiffAutoSuspendTimer(ByVal value As Boolean)
    rCtx.AutoSuspendTimer = value
End Property

'/**
' * @function RiffSuspend
' * @brief Stops the render timer without unloading buffers or releasing WASAPI resources.
' */
Public Sub RiffSuspend()
    RiffStopRenderTimer
End Sub

'/**
' * @function RiffWake
' * @brief Restarts the render timer when the engine is initialized and playback needs to continue.
' * @return {Boolean} True when the timer is running or was started successfully.
' */
Public Function RiffWake() As Boolean
    RiffWake = RiffEnsureRenderTimer()
End Function

'/**
' * @property RiffMaxVoices
' * @brief Returns the maximum number of simultaneous voice slots.
' */
Public Property Get RiffMaxVoices() As Long
    RiffMaxVoices = RIFF_VOICE_COUNT
End Property

'/**
' * @property RiffMaxBuffers
' * @brief Returns the maximum number of static audio buffer slots.
' */
Public Property Get RiffMaxBuffers() As Long
    RiffMaxBuffers = RIFF_BUFFER_COUNT
End Property

'/**
' * @property RiffMaxBuses
' * @brief Returns the maximum number of audio bus slots.
' */
Public Property Get RiffMaxBuses() As Long
    RiffMaxBuses = RIFF_BUS_COUNT
End Property

'/**
' * @property RiffMasterVolume
' * @brief Global master volume for the engine (0.0 to 1.0).
' */
Public Property Get RiffMasterVolume() As Single
    If Not RiffRequireInitialized() Then
        Exit Property
    End If
    
    RiffMasterVolume = rCtx.MasterVolume
End Property
Public Property Let RiffMasterVolume(ByVal value As Single)
    If Not RiffRequireInitialized() Then
        Exit Property
    End If
    
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_UNITY_GAIN Then
        value = RIFF_UNITY_GAIN
    End If
    
    rCtx.MasterVolume = value
End Property

'/**
' * @property RiffBusVolume
' * @brief Gets or sets the volume multiplier for a mixer bus.
' * @param busID The target bus index.
' */
Public Property Get RiffBusVolume(ByVal busID As RiffBusId) As Single
    If Not RiffRequireInitialized() Then
        Exit Property
    End If

    busID = RiffClampBusId(busID)
    RiffBusVolume = rCtx.Buses(busID).Volume
End Property

'/**
' * @property RiffBusVolume
' * @brief Sets the volume multiplier for a mixer bus immediately and cancels any bus fade in progress.
' * @param busID The target bus index.
' * @param value Volume multiplier clamped to 0.0 through RIFF_MAX_BUS_VOLUME.
' */
Public Property Let RiffBusVolume(ByVal busID As RiffBusId, ByVal value As Single)
    If Not RiffRequireInitialized() Then
        Exit Property
    End If

    busID = RiffClampBusId(busID)
    value = RiffClampBusVolume(value)
    
    rCtx.Buses(busID).Volume = value
    rCtx.Buses(busID).targetVolume = value
    rCtx.Buses(busID).FadeStartVolume = value
    rCtx.Buses(busID).FadeFramesCurrent = 0
    rCtx.Buses(busID).FadeFramesTotal = 0
End Property

'/**
' * @property RiffBusMuted
' * @brief Gets or sets whether a mixer bus is muted without changing its stored volume.
' * @param busID The target bus index.
' */
Public Property Get RiffBusMuted(ByVal busID As RiffBusId) As Boolean
    If Not RiffRequireInitialized() Then
        Exit Property
    End If

    busID = RiffClampBusId(busID)
    RiffBusMuted = rCtx.Buses(busID).Muted
End Property

'/**
' * @property RiffBusMuted
' * @brief Mutes or unmutes a mixer bus without changing its stored volume.
' * @param busID The target bus index.
' * @param value True to mute the bus, False to unmute it.
' */
Public Property Let RiffBusMuted(ByVal busID As RiffBusId, ByVal value As Boolean)
    If Not RiffRequireInitialized() Then
        Exit Property
    End If

    busID = RiffClampBusId(busID)
    rCtx.Buses(busID).Muted = value
End Property

'/**
' * @property RiffBusSolo
' * @brief Gets or sets whether a mixer bus is soloed.
' * @param busID The target bus index.
' */
Public Property Get RiffBusSolo(ByVal busID As RiffBusId) As Boolean
    If Not RiffRequireInitialized() Then
        Exit Property
    End If

    busID = RiffClampBusId(busID)
    RiffBusSolo = rCtx.Buses(busID).Solo
End Property

'/**
' * @property RiffBusSolo
' * @brief Solos or unsolos a mixer bus and updates the global solo routing state.
' * @param busID The target bus index.
' * @param value True to solo the bus, False to unsolo it.
' */
Public Property Let RiffBusSolo(ByVal busID As RiffBusId, ByVal value As Boolean)
    If Not RiffRequireInitialized() Then
        Exit Property
    End If

    busID = RiffClampBusId(busID)
    rCtx.Buses(busID).Solo = value
    RiffRefreshSoloState
End Property

'/**
' * @function RiffBusFadeTo
' * @brief Smoothly fades a mixer bus to the requested volume over a duration in milliseconds.
' * @param busID The target bus index.
' * @param targetVolume Destination volume multiplier.
' * @param durationMs Fade duration in milliseconds. Values at or below zero apply immediately.
' */
Public Sub RiffBusFadeTo(ByVal busID As RiffBusId, ByVal targetVolume As Single, Optional ByVal durationMs As Long = 250)
    If Not RiffRequireInitialized() Then
        Exit Sub
    End If

    busID = RiffClampBusId(busID)
    targetVolume = RiffClampBusVolume(targetVolume)

    If durationMs <= 0 Or rCtx.sampleRate <= 0 Then
        rCtx.Buses(busID).Volume = targetVolume
        rCtx.Buses(busID).targetVolume = targetVolume
        rCtx.Buses(busID).FadeStartVolume = targetVolume
        rCtx.Buses(busID).FadeFramesCurrent = 0
        rCtx.Buses(busID).FadeFramesTotal = 0
        Exit Sub
    End If

    rCtx.Buses(busID).FadeStartVolume = rCtx.Buses(busID).Volume
    rCtx.Buses(busID).targetVolume = targetVolume
    rCtx.Buses(busID).FadeFramesCurrent = 0
    rCtx.Buses(busID).FadeFramesTotal = CLng((CDbl(durationMs) * CDbl(rCtx.sampleRate)) / CDbl(RIFF_MS_PER_SEC))
    If rCtx.Buses(busID).FadeFramesTotal < 1 Then
        rCtx.Buses(busID).FadeFramesTotal = 1
    End If
End Sub

'/**
' * @function RiffBusGetPeak
' * @brief Retrieves the current peak amplitude measured after bus routing for metering.
' * @param busID The target bus index.
' * @param peakLeft Variable to receive the left channel peak.
' * @param peakRight Variable to receive the right channel peak.
' */
Public Sub RiffBusGetPeak(ByVal busID As RiffBusId, ByRef peakLeft As Single, ByRef peakRight As Single)
    If Not RiffRequireInitialized() Then
        peakLeft = 0!
        peakRight = 0!
        Exit Sub
    End If

    busID = RiffClampBusId(busID)
    peakLeft = rCtx.Buses(busID).PeakL
    peakRight = rCtx.Buses(busID).PeakR
End Sub

'/**
' * @function RiffBusReset
' * @brief Resets a mixer bus to neutral volume, no fade, no mute, no solo, and cleared peak meters.
' * @param busID The target bus index.
' */
Public Sub RiffBusReset(ByVal busID As RiffBusId)
    If Not RiffRequireInitialized() Then
        Exit Sub
    End If

    busID = RiffClampBusId(busID)
    RiffResetBusState busID
    RiffRefreshSoloState
End Sub

'/**
' * @function RiffMasterGetPeak
' * @brief Retrieves the absolute peak amplitude of the master output for VU Meters.
' * @param peakLeft Variable to store the left channel peak.
' * @param peakRight Variable to store the right channel peak.
' */
Public Sub RiffMasterGetPeak(ByRef peakLeft As Single, ByRef peakRight As Single)
    If Not RiffRequireInitialized() Then
        peakLeft = 0!
        peakRight = 0!
        Exit Sub
    End If
    
    peakLeft = rCtx.MasterPeakL
    peakRight = rCtx.MasterPeakR
End Sub

'/**
' * @function RiffLoad
' * @brief Decodes an audio file from disk into memory for zero-latency playback.
' * @param filePath Full path to the audio file (WAV, MP3, etc.).
' * @return {Long} Buffer handle (0-63), or -1 if failed.
' */
Public Function RiffLoad(ByVal filePath As String) As Long
    RiffLoad = -1
    RiffSetLastError RiffErrorNone

    If Not RiffRequireInitialized() Then
        Exit Function
    End If
    
    Dim slot As Long
    Dim i As Long
    slot = -1
    
    For i = 0 To RIFF_MAX_BUFFER_INDEX
        If Not rCtx.Buffers(i).Active Then
            slot = i
            Exit For
        End If
    Next i
    
    If slot = -1 Then
        RiffSetLastError RiffErrorNoFreeBuffer
        Exit Function
    End If

    If Dir(filePath) = "" Then
        RiffSetLastError RiffErrorFileNotFound
        Exit Function
    End If
    
    #If VBA7 Then
        Dim pReader As LongPtr
        Dim pPartialType As LongPtr
        Dim pAttributes As LongPtr
        Dim pNullPtr As LongPtr
        pNullPtr = 0
    #Else
        Dim pReader As Long
        Dim pPartialType As Long
        Dim pAttributes As Long
        Dim pNullPtr As Long
        pNullPtr = 0
    #End If
    
    Dim hr As Long
    Dim guidResample As GUID
    IIDFromString StrPtr(RIFF_GUID_MF_SOURCE_READER_ENABLE_ADVANCED_VIDEO_PROCESSING), guidResample
    
    hr = MFCreateAttributes(pAttributes, RIFF_MF_ATTRIBUTES_CAPACITY)

    If hr = 0 And pAttributes <> 0 Then
        vCall pAttributes, VTI_MF_ATTRIBUTES_SET_UINT32, VarPtr(guidResample), RIFF_MF_STREAM_SELECT
        hr = MFCreateSourceReaderFromURL(StrPtr(filePath), pAttributes, pReader)
        vCall0 pAttributes, VTI_IUNKNOWN_RELEASE
    Else
        hr = MFCreateSourceReaderFromURL(StrPtr(filePath), 0, pReader)
    End If

    If hr <> 0 Or pReader = 0 Then
        RiffSetLastError RiffErrorComFailure
        Exit Function
    End If

    RiffLoad = CoreProcessSourceReader(pReader, slot)
    If RiffLoad = -1 And rLastError = RiffErrorNone Then
        RiffSetLastError RiffErrorDecodeFailed
    End If
End Function

'/**
' * @function RiffLoadFromMemory
' * @brief Decodes audio data directly from a Byte Array without requiring disk I/O.
' * @param audioData Byte array containing the binary audio file.
' * @return {Long} Buffer handle (0-63), or -1 if failed.
' */
Public Function RiffLoadFromMemory(ByRef audioData() As Byte) As Long
    RiffLoadFromMemory = -1
    RiffSetLastError RiffErrorNone

    If Not RiffRequireInitialized() Then
        Exit Function
    End If
    
    Dim slot As Long
    Dim i As Long
    slot = -1
    
    For i = 0 To RIFF_MAX_BUFFER_INDEX
        If Not rCtx.Buffers(i).Active Then
            slot = i
            Exit For
        End If
    Next i
    
    If slot = -1 Then
        RiffSetLastError RiffErrorNoFreeBuffer
        Exit Function
    End If
    
    #If VBA7 Then
        Dim pStream As LongPtr
        Dim pByteStream As LongPtr
        Dim pReader As LongPtr
        Dim pAttributes As LongPtr
        Dim pAudioBytes As LongPtr
    #Else
        Dim pStream As Long
        Dim pByteStream As Long
        Dim pReader As Long
        Dim pAttributes As Long
        Dim pAudioBytes As Long
    #End If

    Dim cbSize As Long
    If Not RiffTryGetByteArrayData(audioData, pAudioBytes, cbSize) Then
        RiffSetLastError RiffErrorInvalidArgument
        Exit Function
    End If

    pStream = SHCreateMemStream(pAudioBytes, cbSize)
    If pStream = 0 Then
        RiffSetLastError RiffErrorComFailure
        Exit Function
    End If
    
    Dim hr As Long
    hr = MFCreateMFByteStreamOnStream(pStream, pByteStream)

    If hr <> 0 Or pByteStream = 0 Then
        vCall0 pStream, VTI_IUNKNOWN_RELEASE
        RiffSetLastError RiffErrorComFailure
        Exit Function
    End If
    
    Dim guidResample As GUID
    IIDFromString StrPtr(RIFF_GUID_MF_SOURCE_READER_ENABLE_ADVANCED_VIDEO_PROCESSING), guidResample
    
    hr = MFCreateAttributes(pAttributes, RIFF_MF_ATTRIBUTES_CAPACITY)

    If hr = 0 And pAttributes <> 0 Then
        vCall pAttributes, VTI_MF_ATTRIBUTES_SET_UINT32, VarPtr(guidResample), RIFF_MF_STREAM_SELECT
        hr = MFCreateSourceReaderFromByteStream(pByteStream, pAttributes, pReader)
        vCall0 pAttributes, VTI_IUNKNOWN_RELEASE
    Else
        hr = MFCreateSourceReaderFromByteStream(pByteStream, 0, pReader)
    End If
    
    If hr = 0 And pReader <> 0 Then
        RiffLoadFromMemory = CoreProcessSourceReader(pReader, slot)
        If RiffLoadFromMemory = -1 And rLastError = RiffErrorNone Then
            RiffSetLastError RiffErrorDecodeFailed
        End If
    Else
        RiffSetLastError RiffErrorComFailure
    End If

    If pByteStream <> 0 Then
        vCall0 pByteStream, VTI_IUNKNOWN_RELEASE
    End If

    If pStream <> 0 Then
        vCall0 pStream, VTI_IUNKNOWN_RELEASE
    End If
End Function

'/**
' * @function CoreProcessSourceReader
' * @brief Private helper to extract uncompressed PCM data from an IMFSourceReader into memory.
' * @param pReader Pointer to the IMFSourceReader.
' * @param slot Target buffer index to store the extracted audio.
' * @return {Long} Buffer handle if successful, or -1 if failed.
' */
#If VBA7 Then
Private Function CoreProcessSourceReader(ByVal pReader As LongPtr, ByVal slot As Long) As Long
#Else
Private Function CoreProcessSourceReader(ByVal pReader As Long, ByVal slot As Long) As Long
#End If
    CoreProcessSourceReader = -1
    Dim hr As Long
    
    #If VBA7 Then
        Dim pPartialType As LongPtr
        Dim pSample As LongPtr
        Dim pBuffer As LongPtr
        Dim pAudioData As LongPtr
        Dim pNullPtr As LongPtr
        Dim tempPtr As LongPtr
        Dim newPtr As LongPtr
        #If Win64 Then
            Dim llTime As LongLong
            pNullPtr = CLngLng(0)
        #Else
            Dim llTime As Currency
            pNullPtr = CLngPtr(0)
        #End If
    #Else
        Dim pPartialType As Long
        Dim pSample As Long
        Dim pBuffer As Long
        Dim pAudioData As Long
        Dim pNullPtr As Long
        Dim tempPtr As Long
        Dim newPtr As Long
        Dim llTime As Currency
        pNullPtr = 0
    #End If
    
    vCall pReader, VTI_MF_SOURCE_READER_SET_STREAM_SELECTION, MF_SOURCE_READER_ALL_STREAMS, RIFF_MF_STREAM_DESELECT
    vCall pReader, VTI_MF_SOURCE_READER_SET_STREAM_SELECTION, MF_SOURCE_READER_FIRST_AUDIO_STREAM, RIFF_MF_STREAM_SELECT
    
    hr = MFCreateMediaType(pPartialType)
    If hr = 0 And pPartialType <> 0 Then
        Dim wfx_cbSize As Integer
        RtlMoveMemory VarPtr(wfx_cbSize), ByVal (rCtx.MixFormatPtr + RIFF_WFX_CB_SIZE_OFFSET), LenB(wfx_cbSize)
        
        hr = MFInitMediaTypeFromWaveFormatEx(pPartialType, rCtx.MixFormatPtr, RIFF_WFX_BASE_SIZE + CLng(wfx_cbSize))
        If hr = 0 Then
            vCall pReader, VTI_MF_SOURCE_READER_SET_CURRENT_MEDIA_TYPE, MF_SOURCE_READER_FIRST_AUDIO_STREAM, pNullPtr, pPartialType
        End If
        vCall0 pPartialType, VTI_IUNKNOWN_RELEASE
    End If
    
    Dim pSz As Long
    pSz = LenB(pReader)

    Dim dwFlags As Long
    Dim cbMax As Long
    Dim cbLen As Long
    
    Dim rArgs(0 To 5) As Variant
    Dim rTypes(0 To 5) As Integer
    Dim cArgs(0) As Variant
    Dim cTypes(0) As Integer
    Dim lArgs(0 To 2) As Variant
    Dim lTypes(0 To 2) As Integer
    Dim vRet As Variant
    Dim hrInvoke As Long
    Dim i As Long
    
    #If VBA7 Then
        Dim rPtrs(0 To 5) As LongPtr
        Dim cPtrs(0) As LongPtr
        Dim lPtrs(0 To 2) As LongPtr
    #Else
        Dim rPtrs(0 To 5) As Long
        Dim cPtrs(0) As Long
        Dim lPtrs(0 To 2) As Long
    #End If

    rArgs(0) = CLng(MF_SOURCE_READER_FIRST_AUDIO_STREAM)
    rArgs(1) = CLng(0)
    rArgs(2) = pNullPtr
    rArgs(3) = VarPtr(dwFlags)
    rArgs(4) = VarPtr(llTime)
    rArgs(5) = VarPtr(pSample)
    
    For i = 0 To 5
        rTypes(i) = VarType(rArgs(i))
        rPtrs(i) = VarPtr(rArgs(i))
    Next i

    cArgs(0) = VarPtr(pBuffer)
    cTypes(0) = VarType(cArgs(0))
    cPtrs(0) = VarPtr(cArgs(0))

    lArgs(0) = VarPtr(pAudioData)
    lArgs(1) = VarPtr(cbMax)
    lArgs(2) = VarPtr(cbLen)
    
    For i = 0 To 2
        lTypes(i) = VarType(lArgs(i))
        lPtrs(i) = VarPtr(lArgs(i))
    Next i

    Dim totalSize As Long
    Dim currentCap As Long
    
    currentCap = RIFF_DECODE_INITIAL_CAPACITY
    tempPtr = VirtualAlloc(0, currentCap, MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)
    
    If tempPtr = 0 Then
        vCall0 pReader, VTI_IUNKNOWN_RELEASE
        RiffSetLastError RiffErrorMemoryAllocation
        Exit Function
    End If
    
    totalSize = 0
    Do
        pSample = 0
        pBuffer = 0
        pAudioData = 0
        cbMax = 0
        cbLen = 0
        hrInvoke = DispCallFunc(pReader, VTI_MF_SOURCE_READER_READ_SAMPLE * pSz, CC_STDCALL, vbLong, 6, rTypes(0), rPtrs(0), vRet)
        
        If hrInvoke <> 0 Then
            hr = hrInvoke
        Else
            hr = CLng(vRet)
        End If
        
        If hr <> 0 Or (dwFlags And MF_SOURCE_READERF_ENDOFSTREAM) <> 0 Then
            Exit Do
        End If
        
        If pSample <> 0 Then
            hr = vCall(pSample, VTI_MF_SAMPLE_CONVERT_TO_CONTIGUOUS_BUFFER, VarPtr(pBuffer))

            If hr = 0 And pBuffer <> 0 Then
                hr = vCall(pBuffer, VTI_MF_MEDIA_BUFFER_LOCK, VarPtr(pAudioData), VarPtr(cbMax), VarPtr(cbLen))
                
                If hr = 0 And pAudioData <> 0 And cbLen > 0 Then
                    If totalSize + cbLen > currentCap Then
                        currentCap = (totalSize + cbLen) * RIFF_DECODE_GROWTH_FACTOR
                        newPtr = VirtualAlloc(0, currentCap, MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)
                        
                        If newPtr = 0 Then
                            vCall0 pBuffer, VTI_MF_MEDIA_BUFFER_UNLOCK
                            vCall0 pBuffer, VTI_IUNKNOWN_RELEASE
                            vCall0 pSample, VTI_IUNKNOWN_RELEASE
                            vCall0 pReader, VTI_IUNKNOWN_RELEASE
                            If tempPtr <> 0 Then
                                VirtualFree tempPtr, 0, MEM_RELEASE
                            End If
                            RiffSetLastError RiffErrorMemoryAllocation
                            Exit Function
                        End If

                        RtlMoveMemory ByVal newPtr, ByVal tempPtr, totalSize
                        VirtualFree tempPtr, 0, MEM_RELEASE
                        tempPtr = newPtr
                    End If
                    
                    If tempPtr <> 0 Then
                        RtlMoveMemory ByVal (tempPtr + totalSize), ByVal pAudioData, cbLen
                        totalSize = totalSize + cbLen
                    End If
                    
                    vCall0 pBuffer, VTI_MF_MEDIA_BUFFER_UNLOCK
                End If
            End If

            If pBuffer <> 0 Then
                vCall0 pBuffer, VTI_IUNKNOWN_RELEASE
            End If

            vCall0 pSample, VTI_IUNKNOWN_RELEASE
        End If
    Loop

    vCall0 pReader, VTI_IUNKNOWN_RELEASE
    
    If totalSize > 0 Then
        rCtx.Buffers(slot).BufferPtr = VirtualAlloc(0, totalSize, MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)
        
        If rCtx.Buffers(slot).BufferPtr <> 0 Then
            RtlMoveMemory ByVal rCtx.Buffers(slot).BufferPtr, ByVal tempPtr, totalSize
            rCtx.Buffers(slot).BufferLen = totalSize
            rCtx.Buffers(slot).Active = True
            CoreProcessSourceReader = slot
        Else
            RiffSetLastError RiffErrorMemoryAllocation
        End If
    ElseIf rLastError = RiffErrorNone Then
        RiffSetLastError RiffErrorDecodeFailed
    End If
    
    If tempPtr <> 0 Then
        VirtualFree tempPtr, 0, MEM_RELEASE
    End If
End Function

'/**
' * @function RiffUnload
' * @brief Removes a static audio buffer from memory and stops associated playing voices.
' * @param bufferHandle The handle of the buffer to release.
' */
Public Sub RiffUnload(ByVal bufferHandle As Long)
    If Not RiffRequireBufferHandle(bufferHandle) Then
        Exit Sub
    End If
    If Not rCtx.Buffers(bufferHandle).Active Then
        RiffSetLastError RiffErrorInvalidBuffer
        Exit Sub
    End If
    
    Dim i As Long
    For i = 0 To RIFF_MAX_VOICE_INDEX
        If rVoices(i).Active And rVoices(i).BufferIndex = bufferHandle Then
            rVoices(i).Active = False
        End If
    Next i
    
    If rCtx.Buffers(bufferHandle).BufferPtr <> 0 Then
        VirtualFree rCtx.Buffers(bufferHandle).BufferPtr, 0, MEM_RELEASE
        rCtx.Buffers(bufferHandle).BufferPtr = 0
    End If
    
    rCtx.Buffers(bufferHandle).Active = False
End Sub

'/**
' * @property RiffBufferDurationSec
' * @brief Calculates the total duration of a statically loaded buffer in seconds.
' * @param bufferHandle The target buffer.
' */
Public Property Get RiffBufferDurationSec(ByVal bufferHandle As Long) As Single
    If Not RiffRequireBufferHandle(bufferHandle) Then
        Exit Property
    End If
    If Not rCtx.Buffers(bufferHandle).Active Or rCtx.AvgBytesPerSec = 0 Then
        RiffSetLastError RiffErrorInvalidBuffer
        Exit Property
    End If
    
    RiffBufferDurationSec = CSng(rCtx.Buffers(bufferHandle).BufferLen) / CSng(rCtx.AvgBytesPerSec)

End Property

'/**
' * @function RiffExportBufferWav
' * @brief Exports a loaded buffer as a standard 16-bit stereo PCM WAV file.
' * @param bufferHandle Loaded buffer handle.
' * @param filePath Target WAV path.
' * @return {Boolean} True when the file was written successfully.
' */
Public Function RiffExportBufferWav(ByVal bufferHandle As Long, ByVal filePath As String) As Boolean
    RiffSetLastError RiffErrorNone

    If Not RiffRequireBufferHandle(bufferHandle) Then
        Exit Function
    End If
    If Not rCtx.Buffers(bufferHandle).Active Then
        RiffSetLastError RiffErrorInvalidBuffer
        Exit Function
    End If
    If LenB(filePath) = 0 Then
        RiffSetLastError RiffErrorInvalidArgument
        Exit Function
    End If

    Dim nChannels As Integer
    Dim nBlockAlign As Integer
    Dim wBits As Integer
    Dim sampleRate As Long
    Dim frames As Long
    Dim dataBytes As Long
    Dim outBytes() As Byte
    Dim frame As Long
    Dim outIndex As Long
    Dim sL As Single
    Dim sR As Single
    Dim iL As Integer
    Dim iR As Integer
    Dim isFloat As Boolean

    RtlMoveMemory VarPtr(nChannels), ByVal (rCtx.MixFormatPtr + RIFF_WFX_CHANNELS_OFFSET), LenB(nChannels)
    RtlMoveMemory VarPtr(sampleRate), ByVal (rCtx.MixFormatPtr + RIFF_WFX_SAMPLE_RATE_OFFSET), LenB(sampleRate)
    RtlMoveMemory VarPtr(nBlockAlign), ByVal (rCtx.MixFormatPtr + RIFF_WFX_BLOCK_ALIGN_OFFSET), LenB(nBlockAlign)
    RtlMoveMemory VarPtr(wBits), ByVal (rCtx.MixFormatPtr + RIFF_WFX_BITS_OFFSET), LenB(wBits)

    If nChannels <= 0 Or nBlockAlign <= 0 Or sampleRate <= 0 Then
        RiffSetLastError RiffErrorUnsupportedFormat
        Exit Function
    End If

    frames = rCtx.Buffers(bufferHandle).BufferLen \ CLng(nBlockAlign)
    If frames <= 0 Then
        RiffSetLastError RiffErrorInvalidBuffer
        Exit Function
    End If

    isFloat = RiffMixFormatIsFloat32()
    dataBytes = frames * RIFF_WAV_EXPORT_FRAME_BYTES
    ReDim outBytes(0 To dataBytes - 1)

    For frame = 0 To frames - 1
        sL = RiffReadInterleavedSample(rCtx.Buffers(bufferHandle).BufferPtr, frame, 0, nChannels, nBlockAlign, wBits, isFloat)
        If nChannels > 1 Then
            sR = RiffReadInterleavedSample(rCtx.Buffers(bufferHandle).BufferPtr, frame, 1, nChannels, nBlockAlign, wBits, isFloat)
        Else
            sR = sL
        End If

        iL = RiffFloatToPcm16(sL)
        iR = RiffFloatToPcm16(sR)
        outIndex = frame * RIFF_WAV_EXPORT_FRAME_BYTES
        RtlMoveMemory VarPtr(outBytes(outIndex)), VarPtr(iL), LenB(iL)
        RtlMoveMemory VarPtr(outBytes(outIndex + LenB(iL))), VarPtr(iR), LenB(iR)
    Next frame

    RiffExportBufferWav = RiffWritePcm16StereoWav(filePath, sampleRate, outBytes)
End Function

'/**
' * @function RiffRenderOscillatorWav
' * @brief Renders a band-limited oscillator directly to a 16-bit stereo PCM WAV file.
' * @param waveType Oscillator waveform enum.
' * @param frequencyHz Oscillator frequency in Hz.
' * @param durationSec Render duration in seconds.
' * @param filePath Target WAV path.
' * @return {Boolean} True when the file was written successfully.
' */
Public Function RiffRenderOscillatorWav(ByVal waveType As RiffWaveType, ByVal frequencyHz As Single, ByVal durationSec As Single, ByVal filePath As String) As Boolean
    RiffSetLastError RiffErrorNone

    If Not RiffRequireInitialized() Then
        Exit Function
    End If
    If durationSec <= 0! Then
        RiffSetLastError RiffErrorInvalidArgument
        Exit Function
    End If
    If frequencyHz < 1! Then
        frequencyHz = RIFF_DEFAULT_OSCILLATOR_HZ
    End If
    waveType = RiffClampWaveType(waveType)

    Dim frames As Long
    Dim outBytes() As Byte
    Dim frame As Long
    Dim outIndex As Long
    Dim phase As Double
    Dim dt As Double
    Dim sample As Single
    Dim pcm As Integer

    frames = CLng(durationSec * rCtx.sampleRate)
    If frames <= 0 Then
        RiffSetLastError RiffErrorInvalidArgument
        Exit Function
    End If

    ReDim outBytes(0 To (frames * RIFF_WAV_EXPORT_FRAME_BYTES) - 1)
    dt = CDbl(frequencyHz) / CDbl(rCtx.sampleRate)

    For frame = 0 To frames - 1
        sample = RiffOscillatorSampleAtPhase(waveType, phase, dt)
        pcm = RiffFloatToPcm16(sample * RIFF_OSCILLATOR_EXPORT_GAIN)
        outIndex = frame * RIFF_WAV_EXPORT_FRAME_BYTES
        RtlMoveMemory VarPtr(outBytes(outIndex)), VarPtr(pcm), LenB(pcm)
        RtlMoveMemory VarPtr(outBytes(outIndex + LenB(pcm))), VarPtr(pcm), LenB(pcm)
        phase = phase + dt
        If phase >= RIFF_PHASE_FULL Then
            phase = phase - Int(phase)
        End If
    Next frame

    RiffRenderOscillatorWav = RiffWritePcm16StereoWav(filePath, rCtx.sampleRate, outBytes)
End Function

'/**
' * @function RiffFloatToPcm16
' * @brief Converts a normalized float sample to signed 16-bit PCM.
' */
Private Function RiffFloatToPcm16(ByVal value As Single) As Integer
    If value > RIFF_UNITY_GAIN Then
        value = RIFF_UNITY_GAIN
    ElseIf value < RIFF_NEGATIVE_UNITY_GAIN Then
        value = RIFF_NEGATIVE_UNITY_GAIN
    End If

    If value >= 0! Then
        RiffFloatToPcm16 = CInt(value * CSng(RIFF_PCM16_MAX))
    Else
        RiffFloatToPcm16 = CInt(value * CSng(RIFF_PCM16_MIN_MAGNITUDE))
    End If
End Function

'/**
' * @function RiffReadInterleavedSample
' * @brief Reads a sample from an interleaved PCM/float buffer and normalizes it to -1.0..1.0.
' */
#If VBA7 Then
Private Function RiffReadInterleavedSample(ByVal basePtr As LongPtr, ByVal frameIndex As Long, ByVal channelIndex As Long, ByVal nChannels As Integer, ByVal nBlockAlign As Integer, ByVal wBits As Integer, ByVal isFloat As Boolean) As Single
#Else
Private Function RiffReadInterleavedSample(ByVal basePtr As Long, ByVal frameIndex As Long, ByVal channelIndex As Long, ByVal nChannels As Integer, ByVal nBlockAlign As Integer, ByVal wBits As Integer, ByVal isFloat As Boolean) As Single
#End If
    #If VBA7 Then
        Dim pSample As LongPtr
    #Else
        Dim pSample As Long
    #End If

    pSample = basePtr + (CLng(frameIndex) * CLng(nBlockAlign)) + ((CLng(channelIndex) Mod CLng(nChannels)) * (CLng(wBits) \ 8))

    If wBits = RIFF_FLOAT32_BITS And isFloat Then
        Dim f As Single
        RtlMoveMemory VarPtr(f), ByVal pSample, LenB(f)
        RiffReadInterleavedSample = f
    ElseIf wBits = RIFF_FLOAT32_BITS Then
        Dim l As Long
        RtlMoveMemory VarPtr(l), ByVal pSample, LenB(l)
        RiffReadInterleavedSample = CSng(CDbl(l) / RIFF_PCM32_SCALE)
    ElseIf wBits = RIFF_PCM24_BITS Then
        Dim b0 As Byte
        Dim b1 As Byte
        Dim b2 As Byte
        Dim v As Long
        RtlMoveMemory VarPtr(b0), ByVal pSample, LenB(b0)
        RtlMoveMemory VarPtr(b1), ByVal (pSample + RIFF_BYTE_BYTES), LenB(b1)
        RtlMoveMemory VarPtr(b2), ByVal (pSample + (RIFF_BYTE_BYTES * 2)), LenB(b2)
        v = CLng(b0) Or (CLng(b1) * RIFF_BYTE_SHIFT_8) Or (CLng(b2) * RIFF_BYTE_SHIFT_16)
        If (b2 And RIFF_PCM24_SIGN_BIT) <> 0 Then
            v = v Or RIFF_PCM24_SIGN_EXTEND
        End If
        RiffReadInterleavedSample = CSng(CDbl(v) / RIFF_PCM24_SCALE)
    ElseIf wBits = RIFF_PCM16_BITS Then
        Dim i As Integer
        RtlMoveMemory VarPtr(i), ByVal pSample, LenB(i)
        RiffReadInterleavedSample = CSng(CDbl(i) / CDbl(RIFF_PCM16_MIN_MAGNITUDE))
    End If
End Function

'/**
' * @function RiffWritePcm16StereoWav
' * @brief Writes a 16-bit stereo PCM WAV file from an interleaved byte buffer.
' */
Private Function RiffWritePcm16StereoWav(ByVal filePath As String, ByVal sampleRate As Long, ByRef dataBytes() As Byte) As Boolean
    Dim f As Integer
    Dim dataSize As Long
    Dim riffSize As Long
    Dim byteRate As Long
    Dim blockAlign As Integer
    Dim bits As Integer
    Dim channels As Integer
    Dim fmtSize As Long
    Dim audioFormat As Integer

    dataSize = UBound(dataBytes) - LBound(dataBytes) + 1
    riffSize = RIFF_WAV_SIZE_OVERHEAD + dataSize
    channels = RIFF_WAV_EXPORT_CHANNELS
    bits = RIFF_WAV_EXPORT_BITS
    blockAlign = RIFF_WAV_EXPORT_FRAME_BYTES
    byteRate = sampleRate * CLng(blockAlign)
    fmtSize = RIFF_WAV_FMT_CHUNK_SIZE
    audioFormat = RIFF_WAV_FORMAT_PCM

    f = FreeFile
    Open filePath For Binary Access Write As #f
    Put #f, , CByte(Asc("R"))
    Put #f, , CByte(Asc("I"))
    Put #f, , CByte(Asc("F"))
    Put #f, , CByte(Asc("F"))
    Put #f, , riffSize
    Put #f, , CByte(Asc("W"))
    Put #f, , CByte(Asc("A"))
    Put #f, , CByte(Asc("V"))
    Put #f, , CByte(Asc("E"))
    Put #f, , CByte(Asc("f"))
    Put #f, , CByte(Asc("m"))
    Put #f, , CByte(Asc("t"))
    Put #f, , CByte(Asc(" "))
    Put #f, , fmtSize
    Put #f, , audioFormat
    Put #f, , channels
    Put #f, , sampleRate
    Put #f, , byteRate
    Put #f, , blockAlign
    Put #f, , bits
    Put #f, , CByte(Asc("d"))
    Put #f, , CByte(Asc("a"))
    Put #f, , CByte(Asc("t"))
    Put #f, , CByte(Asc("a"))
    Put #f, , dataSize
    Put #f, , dataBytes
    Close #f

    RiffWritePcm16StereoWav = True
End Function

'/**
' * @function RiffOscillatorSampleAtPhase
' * @brief Generates one oscillator sample at a supplied normalized phase using BLEP correction where needed.
' */
Private Function RiffOscillatorSampleAtPhase(ByVal waveType As RiffWaveType, ByVal phase As Double, ByVal dt As Double) As Single
    Dim sample As Double

    Select Case waveType
        Case RiffWaveSine
            sample = Sin(phase * PI2)
        Case RiffWaveSquare
            If phase < RIFF_PHASE_HALF Then
                sample = RIFF_PHASE_FULL
            Else
                sample = -RIFF_PHASE_FULL
            End If
            sample = sample + RiffPolyBLEP(phase, dt)
            Dim t2 As Double
            t2 = phase + RIFF_PHASE_HALF
            If t2 >= RIFF_PHASE_FULL Then
                t2 = t2 - RIFF_PHASE_FULL
            End If
            sample = sample - RiffPolyBLEP(t2, dt)
        Case RiffWaveSawtooth
            sample = (RIFF_BIPOLAR_DOUBLE_SCALE * phase) - RIFF_PHASE_FULL
            sample = sample - RiffPolyBLEP(phase, dt)
        Case Else
            sample = (Rnd() * RIFF_BIPOLAR_SINGLE_SCALE) - RIFF_UNITY_GAIN
    End Select

    RiffOscillatorSampleAtPhase = CSng(sample)
End Function


'/**
' * @function RiffPlay
' * @brief Plays a pre-loaded static audio buffer on an available polyphonic voice channel.
' * @param bufferHandle The target buffer.
' * @return {Long} Voice handle (0-31), or -1 if all voices are occupied.
' */
Public Function RiffPlay(ByVal bufferHandle As Long) As Long
    RiffPlay = RiffPlayBus(bufferHandle, RiffBusMain)
End Function

'/**
' * @function RiffPlayBus
' * @brief Plays a pre-loaded static audio buffer on an available voice and routes it to a bus before activation.
' * @param bufferHandle The target buffer.
' * @param busID Mixer bus that will receive the voice.
' * @return {Long} Voice handle (0-31), or -1 if all voices are occupied.
' */
Public Function RiffPlayBus(ByVal bufferHandle As Long, ByVal busID As RiffBusId) As Long
    RiffPlayBus = -1
    RiffSetLastError RiffErrorNone

    If Not RiffRequireBufferHandle(bufferHandle) Then
        Exit Function
    End If
    If Not rCtx.Buffers(bufferHandle).Active Then
        RiffSetLastError RiffErrorInvalidBuffer
        Exit Function
    End If
    
    busID = RiffClampBusId(busID)
    
    Dim voiceSlot As Long
    voiceSlot = InternalGetFreeVoice()
    
    If voiceSlot = -1 Then
        RiffSetLastError RiffErrorNoFreeVoice
        Exit Function
    End If
    
    InternalResetVoiceDSP voiceSlot
    
    rVoices(voiceSlot).IsOscillator = False
    rVoices(voiceSlot).BufferIndex = bufferHandle
    rVoices(voiceSlot).Position = 0#
    rVoices(voiceSlot).loopEnd = CDbl(rCtx.Buffers(bufferHandle).BufferLen)
    rVoices(voiceSlot).busID = busID

    If Not RiffEnsureRenderTimer() Then
        rVoices(voiceSlot).Playing = False
        rVoices(voiceSlot).Active = False
        RiffPlayBus = -1
        Exit Function
    End If

    rVoices(voiceSlot).Playing = True
    rVoices(voiceSlot).Active = True
    rCtx.IdleTimerTicks = 0
    
    RiffPlayBus = voiceSlot
End Function

'/**
' * @function RiffPlayOscillator
' * @brief Generates and plays a synthesized waveform instantly via math.
' * @param waveType Oscillator waveform enum.
' * @param frequencyHz Pitch frequency in Hz.
' * @return {Long} Voice handle (0-31), or -1 if all voices are occupied.
' */
Public Function RiffPlayOscillator(ByVal waveType As RiffWaveType, ByVal frequencyHz As Single) As Long
    RiffPlayOscillator = -1
    RiffSetLastError RiffErrorNone

    If Not RiffRequireInitialized() Then
        Exit Function
    End If
    
    Dim voiceSlot As Long
    voiceSlot = InternalGetFreeVoice()
    
    If voiceSlot = -1 Then
        RiffSetLastError RiffErrorNoFreeVoice
        Exit Function
    End If
    
    InternalResetVoiceDSP voiceSlot
    
    waveType = RiffClampWaveType(waveType)
    
    If frequencyHz < 1! Then
        frequencyHz = RIFF_DEFAULT_OSCILLATOR_HZ
    End If
    
    rVoices(voiceSlot).IsOscillator = True
    rVoices(voiceSlot).OscType = waveType
    rVoices(voiceSlot).OscFreq = frequencyHz
    rVoices(voiceSlot).OscPhase = 0#
    rVoices(voiceSlot).BufferIndex = -1

    If Not RiffEnsureRenderTimer() Then
        rVoices(voiceSlot).Playing = False
        rVoices(voiceSlot).Active = False
        RiffPlayOscillator = -1
        Exit Function
    End If
    
    rVoices(voiceSlot).Playing = True
    rVoices(voiceSlot).Active = True
    rCtx.IdleTimerTicks = 0
    
    RiffPlayOscillator = voiceSlot
End Function

'/**
' * @function RiffPlayOscillatorBus
' * @brief Generates and plays a synthesized waveform routed to a bus before activation.
' * @param waveType Oscillator waveform enum.
' * @param frequencyHz Pitch frequency in Hz.
' * @param busID Mixer bus that will receive the voice.
' * @return {Long} Voice handle (0-31), or -1 if all voices are occupied.
' */
Public Function RiffPlayOscillatorBus(ByVal waveType As RiffWaveType, ByVal frequencyHz As Single, ByVal busID As RiffBusId) As Long
    RiffPlayOscillatorBus = -1
    RiffSetLastError RiffErrorNone

    If Not RiffRequireInitialized() Then
        Exit Function
    End If

    busID = RiffClampBusId(busID)
    waveType = RiffClampWaveType(waveType)

    If frequencyHz < 1! Then
        frequencyHz = RIFF_DEFAULT_OSCILLATOR_HZ
    End If
    
    Dim voiceSlot As Long
    voiceSlot = InternalGetFreeVoice()
    
    If voiceSlot = -1 Then
        RiffSetLastError RiffErrorNoFreeVoice
        Exit Function
    End If
    
    InternalResetVoiceDSP voiceSlot
    
    rVoices(voiceSlot).IsOscillator = True
    rVoices(voiceSlot).OscType = waveType
    rVoices(voiceSlot).OscFreq = frequencyHz
    rVoices(voiceSlot).OscPhase = 0#
    rVoices(voiceSlot).BufferIndex = -1
    rVoices(voiceSlot).busID = busID

    If Not RiffEnsureRenderTimer() Then
        rVoices(voiceSlot).Playing = False
        rVoices(voiceSlot).Active = False
        RiffPlayOscillatorBus = -1
        Exit Function
    End If

    rVoices(voiceSlot).Playing = True
    rVoices(voiceSlot).Active = True
    rCtx.IdleTimerTicks = 0
    
    RiffPlayOscillatorBus = voiceSlot
End Function

'/**
' * @function InternalGetFreeVoice
' * @brief Scans for an inactive voice channel.
' * @return {Long} Voice index, or -1.
' */
Private Function InternalGetFreeVoice() As Long
    Dim i As Long
    InternalGetFreeVoice = -1
    
    For i = 0 To RIFF_MAX_VOICE_INDEX
        If Not rVoices(i).Active Then
            InternalGetFreeVoice = i
            Exit For
        End If
    Next i
End Function

'/**
' * @function InternalResetVoiceDSP
' * @brief Resets all DSP filters and engine states to neutral values for a voice.
' * @param slot The target voice index.
' */
Private Sub InternalResetVoiceDSP(ByVal slot As Long)
    rVoices(slot).Volume = RIFF_DEFAULT_VOICE_VOLUME
    rVoices(slot).Pitch = RIFF_DEFAULT_VOICE_PITCH
    rVoices(slot).Pan = 0!
    rVoices(slot).Paused = False
    rVoices(slot).busID = RiffBusMain
    rVoices(slot).PeakL = 0!
    rVoices(slot).PeakR = 0!

    rVoices(slot).Distortion = RIFF_DEFAULT_DISTORTION
    rVoices(slot).lowPass = RIFF_DEFAULT_FILTER_CONTROL
    rVoices(slot).highPass = 0!
    rVoices(slot).FilterStateL = 0!
    rVoices(slot).FilterStateR = 0!
    rVoices(slot).FilterStateHP_L = 0!
    rVoices(slot).FilterStateHP_R = 0!
    rVoices(slot).StereoWidth = RIFF_DEFAULT_FILTER_CONTROL

    rVoices(slot).EqBass = RIFF_UNITY_GAIN
    rVoices(slot).EqMid = RIFF_UNITY_GAIN
    rVoices(slot).EqTreble = RIFF_UNITY_GAIN
    rVoices(slot).EqStateLowL = 0!
    rVoices(slot).EqStateLowR = 0!
    rVoices(slot).EqStateHighL = 0!
    rVoices(slot).EqStateHighR = 0!

    rVoices(slot).CompThreshold = RIFF_UNITY_GAIN
    rVoices(slot).CompRatio = RIFF_MIN_COMPRESSOR_RATIO
    rVoices(slot).CompEnv = RIFF_COMP_ENV_FLOOR

    rVoices(slot).BitcrushSteps = 0!
    rVoices(slot).BitcrushDownsample = RIFF_DEFAULT_BITCRUSH_DOWNSAMPLE
    rVoices(slot).BitcrushDsCount = 0
    rVoices(slot).BitcrushLastL = 0!
    rVoices(slot).BitcrushLastR = 0!

    rVoices(slot).RingModFreq = 0!
    rVoices(slot).RingModMix = 0!
    rVoices(slot).RingModPhase = 0#

    rVoices(slot).TremoloRate = 0!
    rVoices(slot).TremoloDepth = 0!
    rVoices(slot).TremoloPhase = 0#

    rVoices(slot).AutoPanRate = 0!
    rVoices(slot).AutoPanDepth = 0!
    rVoices(slot).AutoPanPhase = 0#

    rVoices(slot).ChorusDepth = 0!
    rVoices(slot).ChorusRate = RIFF_DEFAULT_CHORUS_RATE
    rVoices(slot).ChorusPhase = 0#

    rVoices(slot).FlangerRate = RIFF_DEFAULT_FLANGER_RATE
    rVoices(slot).FlangerDepth = 0!
    rVoices(slot).FlangerFeedback = 0!
    rVoices(slot).FlangerPhase = 0#

    rVoices(slot).ReverbMix = 0!
    rVoices(slot).ReverbTime = RIFF_DEFAULT_REVERB_TIME
    rVoices(slot).RevTap1 = Int(RIFF_REVERB_TAP1_SEC * rCtx.sampleRate) * RIFF_WAV_EXPORT_CHANNELS
    rVoices(slot).RevTap2 = Int(RIFF_REVERB_TAP2_SEC * rCtx.sampleRate) * RIFF_WAV_EXPORT_CHANNELS
    rVoices(slot).RevTap3 = Int(RIFF_REVERB_TAP3_SEC * rCtx.sampleRate) * RIFF_WAV_EXPORT_CHANNELS
    rVoices(slot).RevTap4 = Int(RIFF_REVERB_TAP4_SEC * rCtx.sampleRate) * RIFF_WAV_EXPORT_CHANNELS
    rVoices(slot).RevDamp1L = 0!
    rVoices(slot).RevDamp1R = 0!
    rVoices(slot).RevDamp2L = 0!
    rVoices(slot).RevDamp2R = 0!
    rVoices(slot).RevDamp3L = 0!
    rVoices(slot).RevDamp3R = 0!
    rVoices(slot).RevDamp4L = 0!
    rVoices(slot).RevDamp4R = 0!

    rVoices(slot).BqLowPassZ1L = 0!
    rVoices(slot).BqLowPassZ2L = 0!
    rVoices(slot).BqLowPassZ1R = 0!
    rVoices(slot).BqLowPassZ2R = 0!
    rVoices(slot).BqHighPassZ1L = 0!
    rVoices(slot).BqHighPassZ2L = 0!
    rVoices(slot).BqHighPassZ1R = 0!
    rVoices(slot).BqHighPassZ2R = 0!
    rVoices(slot).EqBassZ1L = 0!
    rVoices(slot).EqBassZ2L = 0!
    rVoices(slot).EqBassZ1R = 0!
    rVoices(slot).EqBassZ2R = 0!
    rVoices(slot).EqMidZ1L = 0!
    rVoices(slot).EqMidZ2L = 0!
    rVoices(slot).EqMidZ1R = 0!
    rVoices(slot).EqMidZ2R = 0!
    rVoices(slot).EqTrebleZ1L = 0!
    rVoices(slot).EqTrebleZ2L = 0!
    rVoices(slot).EqTrebleZ1R = 0!
    rVoices(slot).EqTrebleZ2R = 0!

    rVoices(slot).DelayTime = 0!
    rVoices(slot).DelayFeedback = 0!
    rVoices(slot).DelayMix = 0!
    rVoices(slot).RingWritePos = 0

    rVoices(slot).Looping = False
    rVoices(slot).loopStart = 0#
    rVoices(slot).fadeState = 0
    rVoices(slot).FadeFramesTotal = 0
    rVoices(slot).FadeFramesCurrent = 0

    RtlZeroMemory VarPtr(rRingBuf(slot * RIFF_RING_SAMPLES_PER_VOICE)), RIFF_RING_SAMPLES_PER_VOICE * RIFF_SINGLE_BYTES
End Sub

'/**
' * @function RiffPause
' * @brief Pauses a specific voice without killing it.
' * @param voiceHandle The active voice.
' */
Public Sub RiffPause(ByVal voiceHandle As Long)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Sub
    End If
    
    If rVoices(voiceHandle).Active Then
        rVoices(voiceHandle).Paused = True
    End If
End Sub

'/**
' * @function RiffResume
' * @brief Resumes a paused voice.
' * @param voiceHandle The paused voice.
' */
Public Sub RiffResume(ByVal voiceHandle As Long)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Sub
    End If
    
    If rVoices(voiceHandle).Active Then
        rVoices(voiceHandle).Paused = False
        RiffEnsureRenderTimer
    End If
End Sub

'/**
' * @function RiffStop
' * @brief Immediately halts and frees a playing voice.
' * @param voiceHandle The active voice.
' */
Public Sub RiffStop(ByVal voiceHandle As Long)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Sub
    End If
    
    rVoices(voiceHandle).Playing = False
    rVoices(voiceHandle).Active = False
End Sub

'/**
' * @function RiffStopAll
' * @brief Instantly stops all playing voices.
' */
Public Sub RiffStopAll()
    If Not RiffRequireInitialized() Then
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To RIFF_MAX_VOICE_INDEX
        rVoices(i).Playing = False
        rVoices(i).Active = False
    Next i

    RiffStopRenderTimer
End Sub

'/**
' * @function RiffFadeIn
' * @brief Smoothly fades in a voice over the specified duration.
' * @param voiceHandle The active voice.
' * @param durationSec The time in seconds to reach full volume.
' */
Public Sub RiffFadeIn(ByVal voiceHandle As Long, ByVal durationSec As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Sub
    End If
    If durationSec <= 0! Then
        RiffSetLastError RiffErrorInvalidArgument
        Exit Sub
    End If
    
    rVoices(voiceHandle).FadeFramesTotal = CLng(durationSec * rCtx.sampleRate)
    rVoices(voiceHandle).FadeFramesCurrent = 0
    rVoices(voiceHandle).fadeState = 1
End Sub

'/**
' * @function RiffFadeOut
' * @brief Smoothly fades out a voice over the specified duration, stopping it at the end.
' * @param voiceHandle The active voice.
' * @param durationSec The time in seconds to reach zero volume.
' */
Public Sub RiffFadeOut(ByVal voiceHandle As Long, ByVal durationSec As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Sub
    End If
    If durationSec <= 0! Then
        RiffSetLastError RiffErrorInvalidArgument
        Exit Sub
    End If
    
    rVoices(voiceHandle).FadeFramesTotal = CLng(durationSec * rCtx.sampleRate)
    rVoices(voiceHandle).FadeFramesCurrent = 0
    rVoices(voiceHandle).fadeState = 2
End Sub

'/**
' * @function RiffSetLoopRegionSec
' * @brief Constrains the playback to loop between start and end seconds.
' * @param voiceHandle The active voice.
' * @param startSec Loop start position in seconds.
' * @param endSec Loop end position in seconds.
' */
Public Sub RiffSetLoopRegionSec(ByVal voiceHandle As Long, ByVal startSec As Single, ByVal endSec As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Sub
    End If
    If rVoices(voiceHandle).IsOscillator Then
        RiffSetLastError RiffErrorInvalidVoice
        Exit Sub
    End If
    
    Dim sByte As Double
    Dim eByte As Double
    
    sByte = startSec * rCtx.AvgBytesPerSec
    eByte = endSec * rCtx.AvgBytesPerSec
    
    If sByte < 0 Then
        sByte = 0
    End If
    
    Dim bufHandle As Long
    bufHandle = rVoices(voiceHandle).BufferIndex
    If bufHandle < 0 Or bufHandle > RIFF_MAX_BUFFER_INDEX Then
        RiffSetLastError RiffErrorInvalidBuffer
        Exit Sub
    End If
    If Not rCtx.Buffers(bufHandle).Active Then
        RiffSetLastError RiffErrorInvalidBuffer
        Exit Sub
    End If
    
    If eByte > rCtx.Buffers(bufHandle).BufferLen Then
        eByte = rCtx.Buffers(bufHandle).BufferLen
    End If
    
    Dim nBlockAlign As Integer
    RtlMoveMemoryToInteger nBlockAlign, ByVal (rCtx.MixFormatPtr + RIFF_WFX_BLOCK_ALIGN_OFFSET), LenB(nBlockAlign)
    
    Dim align As Long
    align = CLng(nBlockAlign)
    
    sByte = CDbl((CLng(sByte) \ align) * align)
    eByte = CDbl((CLng(eByte) \ align) * align)
    
    rVoices(voiceHandle).loopStart = sByte
    rVoices(voiceHandle).loopEnd = eByte
End Sub


'/**
' * @property RiffVoiceIsPlaying
' * @brief Checks if a voice is actively processing audio.
' */
Public Property Get RiffVoiceIsPlaying(ByVal voiceHandle As Long) As Boolean
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceIsPlaying = (rVoices(voiceHandle).Active And rVoices(voiceHandle).Playing)
End Property

'/**
' * @function RiffVoicePlaying
' * @brief Returns whether a voice handle is valid, active, and currently playing.
' * @param voiceHandle The voice handle to inspect.
' * @return {Boolean} True when the voice exists, is active, and is not stopped.
' */
Public Function RiffVoicePlaying(ByVal voiceHandle As Long) As Boolean
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Function
    End If

    RiffVoicePlaying = (rVoices(voiceHandle).Active And rVoices(voiceHandle).Playing)
End Function

'/**
' * @function RiffVoiceActive
' * @brief Returns whether a voice handle is valid and still owns an active voice slot.
' * @param voiceHandle The voice handle to inspect.
' * @return {Boolean} True when the voice exists and has not been released.
' */
Public Function RiffVoiceActive(ByVal voiceHandle As Long) As Boolean
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Function
    End If

    RiffVoiceActive = rVoices(voiceHandle).Active
End Function

'/**
' * @function RiffFindPlayingVoice
' * @brief Finds the first active playing voice using a specific static buffer.
' * @param bufferHandle The static audio buffer to search for.
' * @param busID Optional bus filter. Pass -1 to search all buses.
' * @return {Long} Voice handle, or -1 when the buffer is not currently playing.
' */
Public Function RiffFindPlayingVoice(ByVal bufferHandle As Long, Optional ByVal busID As Long = -1) As Long
    Dim i As Long

    RiffFindPlayingVoice = -1

    If Not RiffRequireBufferHandle(bufferHandle) Then
        Exit Function
    End If

    If busID > RIFF_MAX_BUS_INDEX Then
        busID = CLng(RiffClampBusId(busID))
    End If

    For i = 0 To RIFF_MAX_VOICE_INDEX
        If rVoices(i).Active Then
            If rVoices(i).Playing Then
                If Not rVoices(i).IsOscillator Then
                    If rVoices(i).BufferIndex = bufferHandle Then
                        If busID < 0 Or rVoices(i).busID = busID Then
                            RiffFindPlayingVoice = i
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Function

'/**
' * @function RiffBufferIsPlaying
' * @brief Checks whether a static audio buffer is already playing on any active voice.
' * @param bufferHandle The static audio buffer to inspect.
' * @param busID Optional bus filter. Pass -1 to search all buses.
' * @return {Boolean} True when at least one active playing voice is using the buffer.
' */
Public Function RiffBufferIsPlaying(ByVal bufferHandle As Long, Optional ByVal busID As Long = -1) As Boolean
    RiffBufferIsPlaying = (RiffFindPlayingVoice(bufferHandle, busID) <> -1)
End Function

'/**
' * @function RiffPlayBusOnce
' * @brief Plays a static buffer on a bus only if the same buffer is not already active on that bus.
' * @param bufferHandle The target buffer.
' * @param busID Mixer bus that will receive the voice.
' * @param looped Whether the voice should loop after creation.
' * @return {Long} Existing or newly-created voice handle, or -1 on failure.
' */
Public Function RiffPlayBusOnce(ByVal bufferHandle As Long, ByVal busID As RiffBusId, Optional ByVal looped As Boolean = False) As Long
    Dim existingVoice As Long

    busID = RiffClampBusId(busID)
    existingVoice = RiffFindPlayingVoice(bufferHandle, CLng(busID))

    If existingVoice <> -1 Then
        RiffPlayBusOnce = existingVoice
        Exit Function
    End If

    RiffPlayBusOnce = RiffPlayBus(bufferHandle, busID)

    If RiffPlayBusOnce <> -1 Then
        rVoices(RiffPlayBusOnce).Looping = looped
    End If
End Function

'/**
' * @property RiffVoiceIsPaused
' * @brief Checks if a voice is currently paused.
' */
Public Property Get RiffVoiceIsPaused(ByVal voiceHandle As Long) As Boolean
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceIsPaused = rVoices(voiceHandle).Paused
End Property

'/**
' * @property RiffVoiceBus
' * @brief Determines which Audio Bus this voice routes to.
' */
Public Property Get RiffVoiceBus(ByVal voiceHandle As Long) As RiffBusId
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceBus = rVoices(voiceHandle).busID
End Property
Public Property Let RiffVoiceBus(ByVal voiceHandle As Long, ByVal value As RiffBusId)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If

    value = RiffClampBusId(value)

    rVoices(voiceHandle).busID = value
End Property

'/**
' * @function RiffVoiceGetPeak
' * @brief Retrieves the instantaneous Peak Amplitude for this voice's VU Meters.
' * @param peakLeft Variable to store the left channel peak.
' * @param peakRight Variable to store the right channel peak.
' */
Public Sub RiffVoiceGetPeak(ByVal voiceHandle As Long, ByRef peakLeft As Single, ByRef peakRight As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        peakLeft = 0!
        peakRight = 0!
        Exit Sub
    End If
    
    peakLeft = rVoices(voiceHandle).PeakL
    peakRight = rVoices(voiceHandle).PeakR
End Sub

'/**
' * @property RiffVoiceLoop
' * @brief Defines whether the voice should restart automatically when it reaches the end.
' */
Public Property Get RiffVoiceLoop(ByVal voiceHandle As Long) As Boolean
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceLoop = rVoices(voiceHandle).Looping
End Property
Public Property Let RiffVoiceLoop(ByVal voiceHandle As Long, ByVal value As Boolean)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    rVoices(voiceHandle).Looping = value
End Property

'/**
' * @property RiffVoicePositionSec
' * @brief Current playback position in seconds.
' */
Public Property Get RiffVoicePositionSec(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If rCtx.AvgBytesPerSec = 0 Then
        RiffSetLastError RiffErrorUnsupportedFormat
        Exit Property
    End If
    RiffVoicePositionSec = CSng(rVoices(voiceHandle).Position) / CSng(rCtx.AvgBytesPerSec)
End Property
Public Property Let RiffVoicePositionSec(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If rVoices(voiceHandle).IsOscillator Then
        Exit Property
    End If
    
    Dim posBytes As Double
    posBytes = CDbl(value) * rCtx.AvgBytesPerSec
    
    If posBytes < 0 Then
        posBytes = 0
    End If
    
    Dim bufHandle As Long
    bufHandle = rVoices(voiceHandle).BufferIndex
    If bufHandle < 0 Or bufHandle > RIFF_MAX_BUFFER_INDEX Then
        Exit Property
    End If
    If Not rCtx.Buffers(bufHandle).Active Then
        Exit Property
    End If
    
    If posBytes >= rCtx.Buffers(bufHandle).BufferLen Then
        posBytes = rCtx.Buffers(bufHandle).BufferLen - 1
    End If
    
    rVoices(voiceHandle).Position = posBytes
End Property

'/**
' * @property RiffVoiceVolume
' * @brief Individual volume for this voice (0.0 to 1.0).
' */
Public Property Get RiffVoiceVolume(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceVolume = rVoices(voiceHandle).Volume
End Property
Public Property Let RiffVoiceVolume(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_UNITY_GAIN Then
        value = RIFF_UNITY_GAIN
    End If
    
    rVoices(voiceHandle).Volume = value
End Property

'/**
' * @property RiffVoicePitch
' * @brief Speed/Pitch modifier (1.0 = Normal, 2.0 = Double speed/octave higher).
' */
Public Property Get RiffVoicePitch(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoicePitch = CSng(rVoices(voiceHandle).Pitch)
End Property
Public Property Let RiffVoicePitch(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    
    If value <= RIFF_MIN_VOICE_PITCH Then
        value = RIFF_MIN_VOICE_PITCH
    End If
    
    rVoices(voiceHandle).Pitch = CDbl(value)
End Property

'/**
' * @property RiffVoicePan
' * @brief Stereo panning (-1.0 = Left, 0.0 = Center, 1.0 = Right).
' */
Public Property Get RiffVoicePan(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoicePan = rVoices(voiceHandle).Pan
End Property
Public Property Let RiffVoicePan(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    
    If value < RIFF_NEGATIVE_UNITY_GAIN Then
        value = RIFF_NEGATIVE_UNITY_GAIN
    End If
    If value > RIFF_UNITY_GAIN Then
        value = RIFF_UNITY_GAIN
    End If
    
    rVoices(voiceHandle).Pan = value
End Property


'/**
' * @property RiffVoiceBitDepth
' * @brief Simulates retro console audio by quantizing amplitude bits.
' */
Public Property Get RiffVoiceBitDepth(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    
    If rVoices(voiceHandle).BitcrushSteps = 0! Then
        RiffVoiceBitDepth = RIFF_MAX_BIT_DEPTH
    Else
        RiffVoiceBitDepth = Log(rVoices(voiceHandle).BitcrushSteps) / Log(2)
    End If
End Property
Public Property Let RiffVoiceBitDepth(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    
    If value >= RIFF_MAX_BIT_DEPTH Then
        rVoices(voiceHandle).BitcrushSteps = 0!
        Exit Property
    End If
    If value < RIFF_MIN_BIT_DEPTH Then
        value = RIFF_MIN_BIT_DEPTH
    End If
    
    rVoices(voiceHandle).BitcrushSteps = 2 ^ value
End Property

'/**
' * @property RiffVoiceSampleRateReduction
' * @brief Creates robotic artifacts by holding samples for N frames.
' */
Public Property Get RiffVoiceSampleRateReduction(ByVal voiceHandle As Long) As Long
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceSampleRateReduction = rVoices(voiceHandle).BitcrushDownsample
End Property
Public Property Let RiffVoiceSampleRateReduction(ByVal voiceHandle As Long, ByVal value As Long)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < RIFF_MIN_DOWNSAMPLE_FACTOR Then
        value = RIFF_MIN_DOWNSAMPLE_FACTOR
    End If
    rVoices(voiceHandle).BitcrushDownsample = value
End Property

'/**
' * @property RiffVoiceRingModFreq
' * @brief The frequency in Hz for the Ring Modulator oscillator.
' */
Public Property Get RiffVoiceRingModFreq(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceRingModFreq = rVoices(voiceHandle).RingModFreq
End Property
Public Property Let RiffVoiceRingModFreq(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    rVoices(voiceHandle).RingModFreq = value
End Property

'/**
' * @property RiffVoiceRingModMix
' * @brief The blend amount of the Ring Modulator effect (0.0 to 1.0).
' */
Public Property Get RiffVoiceRingModMix(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceRingModMix = rVoices(voiceHandle).RingModMix
End Property
Public Property Let RiffVoiceRingModMix(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_UNITY_GAIN Then
        value = RIFF_UNITY_GAIN
    End If
    rVoices(voiceHandle).RingModMix = value
End Property

'/**
' * @property RiffVoiceAutoPanRate
' * @brief Speed of the automatic panning LFO in Hz.
' */
Public Property Get RiffVoiceAutoPanRate(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceAutoPanRate = rVoices(voiceHandle).AutoPanRate
End Property
Public Property Let RiffVoiceAutoPanRate(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_MAX_LFO_RATE_HZ Then
        value = RIFF_MAX_LFO_RATE_HZ
    End If
    rVoices(voiceHandle).AutoPanRate = value
End Property

'/**
' * @property RiffVoiceAutoPanDepth
' * @brief Intensity of the auto-panning effect.
' */
Public Property Get RiffVoiceAutoPanDepth(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceAutoPanDepth = rVoices(voiceHandle).AutoPanDepth
End Property
Public Property Let RiffVoiceAutoPanDepth(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_UNITY_GAIN Then
        value = RIFF_UNITY_GAIN
    End If
    rVoices(voiceHandle).AutoPanDepth = value
End Property

'/**
' * @property RiffVoiceEqBass
' * @brief Low frequency shelf gain (1.0 is flat).
' */
Public Property Get RiffVoiceEqBass(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceEqBass = rVoices(voiceHandle).EqBass
End Property
Public Property Let RiffVoiceEqBass(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_MAX_EQ_GAIN Then
        value = RIFF_MAX_EQ_GAIN
    End If
    rVoices(voiceHandle).EqBass = value
End Property

'/**
' * @property RiffVoiceEqMid
' * @brief Mid frequency band gain (1.0 is flat).
' */
Public Property Get RiffVoiceEqMid(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceEqMid = rVoices(voiceHandle).EqMid
End Property
Public Property Let RiffVoiceEqMid(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_MAX_EQ_GAIN Then
        value = RIFF_MAX_EQ_GAIN
    End If
    rVoices(voiceHandle).EqMid = value
End Property

'/**
' * @property RiffVoiceEqTreble
' * @brief High frequency shelf gain (1.0 is flat).
' */
Public Property Get RiffVoiceEqTreble(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceEqTreble = rVoices(voiceHandle).EqTreble
End Property
Public Property Let RiffVoiceEqTreble(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_MAX_EQ_GAIN Then
        value = RIFF_MAX_EQ_GAIN
    End If
    rVoices(voiceHandle).EqTreble = value
End Property

'/**
' * @property RiffVoiceCompressorThreshold
' * @brief Volume level at which the compressor starts reducing gain.
' */
Public Property Get RiffVoiceCompressorThreshold(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceCompressorThreshold = rVoices(voiceHandle).CompThreshold
End Property
Public Property Let RiffVoiceCompressorThreshold(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value <= 0! Then
        value = RIFF_MIN_COMPRESSOR_THRESHOLD
    End If
    If value > RIFF_UNITY_GAIN Then
        value = RIFF_UNITY_GAIN
    End If
    rVoices(voiceHandle).CompThreshold = value
End Property

'/**
' * @property RiffVoiceCompressorRatio
' * @brief Amount of gain reduction applied when signal exceeds the threshold.
' */
Public Property Get RiffVoiceCompressorRatio(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceCompressorRatio = rVoices(voiceHandle).CompRatio
End Property
Public Property Let RiffVoiceCompressorRatio(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < RIFF_MIN_COMPRESSOR_RATIO Then
        value = RIFF_MIN_COMPRESSOR_RATIO
    End If
    If value > RIFF_MAX_COMPRESSOR_RATIO Then
        value = RIFF_MAX_COMPRESSOR_RATIO
    End If
    rVoices(voiceHandle).CompRatio = value
End Property

'/**
' * @property RiffVoiceFlangerDepth
' * @brief Blend amount for the Flanger effect.
' */
Public Property Get RiffVoiceFlangerDepth(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceFlangerDepth = rVoices(voiceHandle).FlangerDepth
End Property
Public Property Let RiffVoiceFlangerDepth(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_UNITY_GAIN Then
        value = RIFF_UNITY_GAIN
    End If
    rVoices(voiceHandle).FlangerDepth = value
End Property

'/**
' * @property RiffVoiceFlangerRate
' * @brief Sweep rate of the Flanger in Hz.
' */
Public Property Get RiffVoiceFlangerRate(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceFlangerRate = rVoices(voiceHandle).FlangerRate
End Property
Public Property Let RiffVoiceFlangerRate(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < RIFF_MIN_MOD_RATE_HZ Then
        value = RIFF_MIN_MOD_RATE_HZ
    End If
    If value > RIFF_MAX_MOD_RATE_HZ Then
        value = RIFF_MAX_MOD_RATE_HZ
    End If
    rVoices(voiceHandle).FlangerRate = value
End Property

'/**
' * @property RiffVoiceFlangerFeedback
' * @brief Resonance intensity of the Flanger effect.
' */
Public Property Get RiffVoiceFlangerFeedback(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceFlangerFeedback = rVoices(voiceHandle).FlangerFeedback
End Property
Public Property Let RiffVoiceFlangerFeedback(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_MAX_FEEDBACK Then
        value = RIFF_MAX_FEEDBACK
    End If
    rVoices(voiceHandle).FlangerFeedback = value
End Property

'/**
' * @property RiffVoiceDistortion
' * @brief Digital clipping multiplier.
' */
Public Property Get RiffVoiceDistortion(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceDistortion = rVoices(voiceHandle).Distortion
End Property
Public Property Let RiffVoiceDistortion(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < RIFF_DEFAULT_DISTORTION Then
        value = RIFF_DEFAULT_DISTORTION
    End If
    rVoices(voiceHandle).Distortion = value
End Property

'/**
' * @property RiffVoiceLowPass
' * @brief Muffles the audio by filtering high frequencies.
' */
Public Property Get RiffVoiceLowPass(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceLowPass = rVoices(voiceHandle).lowPass
End Property
Public Property Let RiffVoiceLowPass(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value <= 0! Then
        value = RIFF_MIN_LOWPASS_CONTROL
    End If
    If value > RIFF_UNITY_GAIN Then
        value = RIFF_UNITY_GAIN
    End If
    rVoices(voiceHandle).lowPass = value
End Property

'/**
' * @property RiffVoiceHighPass
' * @brief Thins out the audio by filtering low frequencies.
' */
Public Property Get RiffVoiceHighPass(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceHighPass = rVoices(voiceHandle).highPass
End Property
Public Property Let RiffVoiceHighPass(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_MAX_STEREO_WIDTH Then
        value = RIFF_MAX_STEREO_WIDTH
    End If
    rVoices(voiceHandle).highPass = value
End Property

'/**
' * @property RiffVoiceStereoWidth
' * @brief Adjusts the perceived width of the stereo field.
' */
Public Property Get RiffVoiceStereoWidth(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceStereoWidth = rVoices(voiceHandle).StereoWidth
End Property
Public Property Let RiffVoiceStereoWidth(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_MAX_STEREO_WIDTH_CONTROL Then
        value = RIFF_MAX_STEREO_WIDTH_CONTROL
    End If
    rVoices(voiceHandle).StereoWidth = value
End Property

'/**
' * @property RiffVoiceTremoloRate
' * @brief Speed of the volume oscillation LFO in Hz.
' */
Public Property Get RiffVoiceTremoloRate(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceTremoloRate = rVoices(voiceHandle).TremoloRate
End Property
Public Property Let RiffVoiceTremoloRate(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_MAX_LFO_RATE_HZ Then
        value = RIFF_MAX_LFO_RATE_HZ
    End If
    rVoices(voiceHandle).TremoloRate = value
End Property

'/**
' * @property RiffVoiceTremoloDepth
' * @brief Intensity of the volume oscillation effect.
' */
Public Property Get RiffVoiceTremoloDepth(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceTremoloDepth = rVoices(voiceHandle).TremoloDepth
End Property
Public Property Let RiffVoiceTremoloDepth(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_UNITY_GAIN Then
        value = RIFF_UNITY_GAIN
    End If
    rVoices(voiceHandle).TremoloDepth = value
End Property

'/**
' * @property RiffVoiceChorusDepth
' * @brief Wet mix amount for the multi-voice Chorus effect.
' */
Public Property Get RiffVoiceChorusDepth(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceChorusDepth = rVoices(voiceHandle).ChorusDepth
End Property
Public Property Let RiffVoiceChorusDepth(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_UNITY_GAIN Then
        value = RIFF_UNITY_GAIN
    End If
    rVoices(voiceHandle).ChorusDepth = value
End Property

'/**
' * @property RiffVoiceChorusRate
' * @brief LFO rate governing the Chorus pitch modulation.
' */
Public Property Get RiffVoiceChorusRate(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceChorusRate = rVoices(voiceHandle).ChorusRate
End Property
Public Property Let RiffVoiceChorusRate(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < RIFF_MIN_MOD_RATE_HZ Then
        value = RIFF_MIN_MOD_RATE_HZ
    End If
    If value > RIFF_MAX_MOD_RATE_HZ Then
        value = RIFF_MAX_MOD_RATE_HZ
    End If
    rVoices(voiceHandle).ChorusRate = value
End Property

'/**
' * @property RiffVoiceReverbMix
' * @brief Blend of the simulated spatial room reverberation.
' */
Public Property Get RiffVoiceReverbMix(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceReverbMix = rVoices(voiceHandle).ReverbMix
End Property
Public Property Let RiffVoiceReverbMix(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_UNITY_GAIN Then
        value = RIFF_UNITY_GAIN
    End If
    rVoices(voiceHandle).ReverbMix = value
End Property

'/**
' * @property RiffVoiceReverbTime
' * @brief Determines the decay length / simulated room size.
' */
Public Property Get RiffVoiceReverbTime(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceReverbTime = rVoices(voiceHandle).ReverbTime
End Property
Public Property Let RiffVoiceReverbTime(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_MAX_FEEDBACK Then
        value = RIFF_MAX_FEEDBACK
    End If
    rVoices(voiceHandle).ReverbTime = value
End Property

'/**
' * @property RiffVoiceDelayTime
' * @brief Interval between consecutive delay echoes.
' */
Public Property Get RiffVoiceDelayTime(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceDelayTime = rVoices(voiceHandle).DelayTime
End Property
Public Property Let RiffVoiceDelayTime(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_UNITY_GAIN Then
        value = RIFF_UNITY_GAIN
    End If
    rVoices(voiceHandle).DelayTime = value
End Property

'/**
' * @property RiffVoiceDelayFeedback
' * @brief The amount of signal fed back into the delay line to create decaying echoes.
' */
Public Property Get RiffVoiceDelayFeedback(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceDelayFeedback = rVoices(voiceHandle).DelayFeedback
End Property
Public Property Let RiffVoiceDelayFeedback(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_MAX_FEEDBACK Then
        value = RIFF_MAX_FEEDBACK
    End If
    rVoices(voiceHandle).DelayFeedback = value
End Property

'/**
' * @property RiffVoiceDelayMix
' * @brief Blend of the Echo/Delay effect into the main output.
' */
Public Property Get RiffVoiceDelayMix(ByVal voiceHandle As Long) As Single
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    RiffVoiceDelayMix = rVoices(voiceHandle).DelayMix
End Property
Public Property Let RiffVoiceDelayMix(ByVal voiceHandle As Long, ByVal value As Single)
    If Not RiffRequireVoiceHandle(voiceHandle) Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > RIFF_UNITY_GAIN Then
        value = RIFF_UNITY_GAIN
    End If
    rVoices(voiceHandle).DelayMix = value
End Property


'/**
' * @function RiffTimerCallback
' * @brief Core multimedia timer callback. Executes the DSP pipeline and writes to WASAPI.
' * @param hWnd Window Handle (not used).
' * @param uMsg System Message (not used).
' * @param idEvent Timer Identifier (not used).
' * @param dwTime System Time (not used).
' */

'/**
' * @function RiffClamp
' * @brief Bounds a floating-point value into a deterministic range for DSP safety.
' */
Private Function RiffClamp(ByVal value As Single, ByVal minValue As Single, ByVal maxValue As Single) As Single
    If value < minValue Then
        RiffClamp = minValue
    ElseIf value > maxValue Then
        RiffClamp = maxValue
    Else
        RiffClamp = value
    End If
End Function

'/**
' * @function RiffEnsureSingleScratch
' * @brief Grows a reusable Single scratch buffer only when the current capacity is too small.
' */
Private Sub RiffEnsureSingleScratch(ByRef buffer() As Single, ByRef capacity As Long, ByVal needed As Long)
    If needed < 1 Then
        needed = 1
    End If
    If capacity < needed Then
        ReDim buffer(0 To needed - 1)
        capacity = needed
    End If
End Sub

'/**
' * @function RiffEnsureLongScratch
' * @brief Grows a reusable Long scratch buffer only when the current capacity is too small.
' */
Private Sub RiffEnsureLongScratch(ByRef buffer() As Long, ByRef capacity As Long, ByVal needed As Long)
    If needed < 1 Then
        needed = 1
    End If
    If capacity < needed Then
        ReDim buffer(0 To needed - 1)
        capacity = needed
    End If
End Sub

'/**
' * @function RiffEnsureIntegerScratch
' * @brief Grows a reusable Integer scratch buffer only when the current capacity is too small.
' */
Private Sub RiffEnsureIntegerScratch(ByRef buffer() As Integer, ByRef capacity As Long, ByVal needed As Long)
    If needed < 1 Then
        needed = 1
    End If
    If capacity < needed Then
        ReDim buffer(0 To needed - 1)
        capacity = needed
    End If
End Sub

'/**
' * @function RiffClearSingleScratch
' * @brief Clears the active region of a reusable Single scratch buffer.
' */
Private Sub RiffClearSingleScratch(ByRef buffer() As Single, ByVal itemCount As Long)
    If itemCount > 0 Then
        RtlZeroMemory VarPtr(buffer(0)), itemCount * RIFF_SINGLE_BYTES
    End If
End Sub

'/**
' * @function RiffClearLongScratch
' * @brief Clears the active region of a reusable Long scratch buffer.
' */
Private Sub RiffClearLongScratch(ByRef buffer() As Long, ByVal itemCount As Long)
    If itemCount > 0 Then
        RtlZeroMemory VarPtr(buffer(0)), itemCount * RIFF_LONG_BYTES
    End If
End Sub

'/**
' * @function RiffClearIntegerScratch
' * @brief Clears the active region of a reusable Integer scratch buffer.
' */
Private Sub RiffClearIntegerScratch(ByRef buffer() As Integer, ByVal itemCount As Long)
    If itemCount > 0 Then
        RtlZeroMemory VarPtr(buffer(0)), itemCount * RIFF_INTEGER_BYTES
    End If
End Sub

'/**
' * @function RiffPolyBLEP
' * @brief Produces a polynomial band-limiting correction at waveform discontinuities.
' */
Private Function RiffPolyBLEP(ByVal t As Double, ByVal dt As Double) As Single
    If dt <= 0# Then
        Exit Function
    End If

    If t < dt Then
        t = t / dt
        RiffPolyBLEP = CSng((t + t) - (t * t) - RIFF_PHASE_FULL)
    ElseIf t > RIFF_PHASE_FULL - dt Then
        t = (t - RIFF_PHASE_FULL) / dt
        RiffPolyBLEP = CSng((t * t) + (t + t) + RIFF_PHASE_FULL)
    Else
        RiffPolyBLEP = 0!
    End If
End Function

'/**
' * @function RiffNextOscillatorSample
' * @brief Generates one oscillator sample using BLEP band-limiting for discontinuous waveforms.
' */
Private Function RiffNextOscillatorSample(ByVal voiceIndex As Long) As Single
    Dim dt As Double
    Dim phase01 As Double
    Dim sample As Single
    Dim edge As Double

    If rCtx.sampleRate <= 0 Then
        Exit Function
    End If

    dt = CDbl(rVoices(voiceIndex).OscFreq) / CDbl(rCtx.sampleRate)
    If dt <= 0# Then
        dt = RIFF_OSCILLATOR_FALLBACK_HZ / CDbl(rCtx.sampleRate)
    End If
    If dt > RIFF_OSCILLATOR_MAX_DT Then
        dt = RIFF_OSCILLATOR_MAX_DT
    End If

    phase01 = rVoices(voiceIndex).OscPhase / PI2
    phase01 = phase01 - Fix(phase01)
    If phase01 < 0# Then
        phase01 = phase01 + RIFF_PHASE_FULL
    End If

    Select Case rVoices(voiceIndex).OscType
        Case RiffWaveSine
            sample = CSng(Sin(rVoices(voiceIndex).OscPhase))
        Case RiffWaveSquare
            If phase01 < RIFF_HALF_SCALE Then
                sample = RIFF_OSCILLATOR_BLEP_LEVEL
            Else
                sample = -RIFF_OSCILLATOR_BLEP_LEVEL
            End If
            sample = sample + RIFF_OSCILLATOR_BLEP_LEVEL * RiffPolyBLEP(phase01, dt)
            edge = phase01 + RIFF_HALF_SCALE
            If edge >= RIFF_PHASE_FULL Then
                edge = edge - RIFF_PHASE_FULL
            End If
            sample = sample - RIFF_OSCILLATOR_BLEP_LEVEL * RiffPolyBLEP(edge, dt)
        Case RiffWaveSawtooth
            sample = CSng((RIFF_BIPOLAR_DOUBLE_SCALE * phase01) - RIFF_PHASE_FULL)
            sample = sample - RiffPolyBLEP(phase01, dt)
            sample = sample * RIFF_OSCILLATOR_BLEP_LEVEL
        Case Else
            sample = (Rnd() * RIFF_BIPOLAR_SINGLE_SCALE) - RIFF_UNITY_GAIN
    End Select

    rVoices(voiceIndex).OscPhase = rVoices(voiceIndex).OscPhase + (dt * PI2)
    If rVoices(voiceIndex).OscPhase >= PI2 Then
        rVoices(voiceIndex).OscPhase = rVoices(voiceIndex).OscPhase - PI2
    End If

    RiffNextOscillatorSample = sample
End Function

'/**
' * @function RiffEngineHasActivePlayback
' * @brief Returns True only when at least one voice can produce audible samples.
' * Prevents the render loop from continuously queuing silence while the engine is idle.
' */
Private Function RiffEngineHasActivePlayback() As Boolean
    Dim i As Long

    For i = 0 To RIFF_MAX_VOICE_INDEX
        If rVoices(i).Active And rVoices(i).Playing And Not rVoices(i).Paused Then
            If rVoices(i).IsOscillator Then
                RiffEngineHasActivePlayback = True
                Exit Function
            End If

            If rVoices(i).BufferIndex >= 0 And rVoices(i).BufferIndex <= RIFF_MAX_BUFFER_INDEX Then
                If rCtx.Buffers(rVoices(i).BufferIndex).Active Then
                    RiffEngineHasActivePlayback = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function


'/**
' * @function RiffBiquadProcess
' * @brief Processes one sample through a transposed direct-form II biquad section.
' */
Private Function RiffBiquadProcess(ByVal inputSample As Single, ByRef z1 As Single, ByRef z2 As Single, ByVal b0 As Single, ByVal b1 As Single, ByVal b2 As Single, ByVal a1 As Single, ByVal a2 As Single) As Single
    Dim y As Single

    y = (b0 * inputSample) + z1
    z1 = (b1 * inputSample) - (a1 * y) + z2
    z2 = (b2 * inputSample) - (a2 * y)
    RiffBiquadProcess = y
End Function

'/**
' * @function RiffBiquadLowPassCoeffs
' * @brief Calculates normalized second-order low-pass filter coefficients.
' */
Private Sub RiffBiquadLowPassCoeffs(ByVal cutoffHz As Single, ByVal q As Single, ByRef b0 As Single, ByRef b1 As Single, ByRef b2 As Single, ByRef a1 As Single, ByRef a2 As Single)
    Dim omega As Double
    Dim sn As Double
    Dim cs As Double
    Dim alpha As Double
    Dim a0 As Double

    cutoffHz = RiffClamp(cutoffHz, RIFF_FILTER_MIN_CUTOFF_HZ, CSng(rCtx.sampleRate) * RIFF_LOWPASS_MAX_SAMPLE_RATE_RATIO)
    q = RiffClamp(q, RIFF_BIQUAD_MIN_Q, RIFF_BIQUAD_MAX_Q)

    omega = PI2 * CDbl(cutoffHz) / CDbl(rCtx.sampleRate)
    sn = Sin(omega)
    cs = Cos(omega)
    alpha = sn / (RIFF_BIPOLAR_DOUBLE_SCALE * CDbl(q))
    a0 = RIFF_PHASE_FULL + alpha

    b0 = CSng(((RIFF_PHASE_FULL - cs) * CDbl(RIFF_HALF_SCALE)) / a0)
    b1 = CSng((RIFF_PHASE_FULL - cs) / a0)
    b2 = b0
    a1 = CSng((-RIFF_BIPOLAR_DOUBLE_SCALE * cs) / a0)
    a2 = CSng((RIFF_PHASE_FULL - alpha) / a0)
End Sub

'/**
' * @function RiffBiquadHighPassCoeffs
' * @brief Calculates normalized second-order high-pass filter coefficients.
' */
Private Sub RiffBiquadHighPassCoeffs(ByVal cutoffHz As Single, ByVal q As Single, ByRef b0 As Single, ByRef b1 As Single, ByRef b2 As Single, ByRef a1 As Single, ByRef a2 As Single)
    Dim omega As Double
    Dim sn As Double
    Dim cs As Double
    Dim alpha As Double
    Dim a0 As Double

    cutoffHz = RiffClamp(cutoffHz, RIFF_FILTER_MIN_CUTOFF_HZ, CSng(rCtx.sampleRate) * RIFF_LOWPASS_MAX_SAMPLE_RATE_RATIO)
    q = RiffClamp(q, RIFF_BIQUAD_MIN_Q, RIFF_BIQUAD_MAX_Q)

    omega = PI2 * CDbl(cutoffHz) / CDbl(rCtx.sampleRate)
    sn = Sin(omega)
    cs = Cos(omega)
    alpha = sn / (RIFF_BIPOLAR_DOUBLE_SCALE * CDbl(q))
    a0 = RIFF_PHASE_FULL + alpha

    b0 = CSng(((RIFF_PHASE_FULL + cs) * CDbl(RIFF_HALF_SCALE)) / a0)
    b1 = CSng((-(RIFF_PHASE_FULL + cs)) / a0)
    b2 = b0
    a1 = CSng((-RIFF_BIPOLAR_DOUBLE_SCALE * cs) / a0)
    a2 = CSng((RIFF_PHASE_FULL - alpha) / a0)
End Sub

'/**
' * @function RiffBiquadPeakCoeffs
' * @brief Calculates normalized parametric EQ coefficients for a single band.
' */
Private Sub RiffBiquadPeakCoeffs(ByVal freqHz As Single, ByVal q As Single, ByVal gain As Single, ByRef b0 As Single, ByRef b1 As Single, ByRef b2 As Single, ByRef a1 As Single, ByRef a2 As Single)
    Dim omega As Double
    Dim sn As Double
    Dim cs As Double
    Dim alpha As Double
    Dim amp As Double
    Dim a0 As Double

    freqHz = RiffClamp(freqHz, RIFF_FILTER_MIN_CUTOFF_HZ, CSng(rCtx.sampleRate) * RIFF_LOWPASS_MAX_SAMPLE_RATE_RATIO)
    q = RiffClamp(q, RIFF_BIQUAD_MIN_Q, RIFF_BIQUAD_MAX_Q)
    gain = RiffClamp(gain, RIFF_BIQUAD_MIN_GAIN, RIFF_BIQUAD_MAX_GAIN)

    omega = PI2 * CDbl(freqHz) / CDbl(rCtx.sampleRate)
    sn = Sin(omega)
    cs = Cos(omega)
    alpha = sn / (RIFF_BIPOLAR_DOUBLE_SCALE * CDbl(q))
    amp = Sqr(CDbl(gain))
    a0 = RIFF_PHASE_FULL + (alpha / amp)

    b0 = CSng((RIFF_PHASE_FULL + (alpha * amp)) / a0)
    b1 = CSng((-RIFF_BIPOLAR_DOUBLE_SCALE * cs) / a0)
    b2 = CSng((RIFF_PHASE_FULL - (alpha * amp)) / a0)
    a1 = CSng((-RIFF_BIPOLAR_DOUBLE_SCALE * cs) / a0)
    a2 = CSng((RIFF_PHASE_FULL - (alpha / amp)) / a0)
End Sub

'/**
' * @function RiffProcessVoiceFilters
' * @brief Applies biquad low-pass, high-pass, and three-band parametric EQ to one stereo sample.
' */
Private Sub RiffProcessVoiceFilters(ByVal voiceIndex As Long, ByRef leftSample As Single, ByRef rightSample As Single, ByVal lowPass As Single, ByVal highPass As Single, ByVal bassGain As Single, ByVal midGain As Single, ByVal trebleGain As Single)
    Static lastVoice As Long
    Static lastLp As Single, lastHp As Single
    Static lastEqB As Single, lastEqM As Single, lastEqT As Single
    
    Static b0_lp As Single, b1_lp As Single, b2_lp As Single, a1_lp As Single, a2_lp As Single
    Static b0_hp As Single, b1_hp As Single, b2_hp As Single, a1_hp As Single, a2_hp As Single
    Static b0_b As Single, b1_b As Single, b2_b As Single, a1_b As Single, a2_b As Single
    Static b0_m As Single, b1_m As Single, b2_m As Single, a1_m As Single, a2_m As Single
    Static b0_t As Single, b1_t As Single, b2_t As Single, a1_t As Single, a2_t As Single
    
    Dim recalc As Boolean
    Dim cutoff As Single
    
    If voiceIndex <> lastVoice Then
        recalc = True
        lastVoice = voiceIndex
    End If
    
    If rCtx.sampleRate <= 0 Then Exit Sub
    
    If lowPass < RIFF_LOWPASS_BYPASS_THRESHOLD Then
        If recalc Or lowPass <> lastLp Then
            cutoff = RIFF_LOWPASS_MIN_CUTOFF_HZ + ((RiffClamp(lowPass, 0!, RIFF_UNITY_GAIN) ^ RIFF_BIPOLAR_SINGLE_SCALE) * ((CSng(rCtx.sampleRate) * RIFF_LOWPASS_MAX_SAMPLE_RATE_RATIO) - RIFF_LOWPASS_MIN_CUTOFF_HZ))
            RiffBiquadLowPassCoeffs cutoff, RIFF_FILTER_DEFAULT_Q, b0_lp, b1_lp, b2_lp, a1_lp, a2_lp
            lastLp = lowPass
        End If
        leftSample = RiffBiquadProcess(leftSample, rVoices(voiceIndex).BqLowPassZ1L, rVoices(voiceIndex).BqLowPassZ2L, b0_lp, b1_lp, b2_lp, a1_lp, a2_lp)
        rightSample = RiffBiquadProcess(rightSample, rVoices(voiceIndex).BqLowPassZ1R, rVoices(voiceIndex).BqLowPassZ2R, b0_lp, b1_lp, b2_lp, a1_lp, a2_lp)
    End If
    
    If highPass > 0! Then
        If recalc Or highPass <> lastHp Then
            cutoff = RIFF_FILTER_MIN_CUTOFF_HZ + ((RiffClamp(highPass, 0!, RIFF_UNITY_GAIN) ^ RIFF_BIPOLAR_SINGLE_SCALE) * ((CSng(rCtx.sampleRate) * RIFF_HIGHPASS_MAX_SAMPLE_RATE_RATIO) - RIFF_FILTER_MIN_CUTOFF_HZ))
            RiffBiquadHighPassCoeffs cutoff, RIFF_FILTER_DEFAULT_Q, b0_hp, b1_hp, b2_hp, a1_hp, a2_hp
            lastHp = highPass
        End If
        leftSample = RiffBiquadProcess(leftSample, rVoices(voiceIndex).BqHighPassZ1L, rVoices(voiceIndex).BqHighPassZ2L, b0_hp, b1_hp, b2_hp, a1_hp, a2_hp)
        rightSample = RiffBiquadProcess(rightSample, rVoices(voiceIndex).BqHighPassZ1R, rVoices(voiceIndex).BqHighPassZ2R, b0_hp, b1_hp, b2_hp, a1_hp, a2_hp)
    End If
    
    If bassGain <> 1! Then
        If recalc Or bassGain <> lastEqB Then
            RiffBiquadPeakCoeffs RIFF_EQ_BASS_CENTER_HZ, RIFF_EQ_DEFAULT_Q, bassGain, b0_b, b1_b, b2_b, a1_b, a2_b
            lastEqB = bassGain
        End If
        leftSample = RiffBiquadProcess(leftSample, rVoices(voiceIndex).EqBassZ1L, rVoices(voiceIndex).EqBassZ2L, b0_b, b1_b, b2_b, a1_b, a2_b)
        rightSample = RiffBiquadProcess(rightSample, rVoices(voiceIndex).EqBassZ1R, rVoices(voiceIndex).EqBassZ2R, b0_b, b1_b, b2_b, a1_b, a2_b)
    End If
    
    If midGain <> 1! Then
        If recalc Or midGain <> lastEqM Then
            RiffBiquadPeakCoeffs RIFF_EQ_MID_CENTER_HZ, RIFF_EQ_MID_Q, midGain, b0_m, b1_m, b2_m, a1_m, a2_m
            lastEqM = midGain
        End If
        leftSample = RiffBiquadProcess(leftSample, rVoices(voiceIndex).EqMidZ1L, rVoices(voiceIndex).EqMidZ2L, b0_m, b1_m, b2_m, a1_m, a2_m)
        rightSample = RiffBiquadProcess(rightSample, rVoices(voiceIndex).EqMidZ1R, rVoices(voiceIndex).EqMidZ2R, b0_m, b1_m, b2_m, a1_m, a2_m)
    End If
    
    If trebleGain <> 1! Then
        If recalc Or trebleGain <> lastEqT Then
            RiffBiquadPeakCoeffs RIFF_EQ_TREBLE_CENTER_HZ, RIFF_EQ_DEFAULT_Q, trebleGain, b0_t, b1_t, b2_t, a1_t, a2_t
            lastEqT = trebleGain
        End If
        leftSample = RiffBiquadProcess(leftSample, rVoices(voiceIndex).EqTrebleZ1L, rVoices(voiceIndex).EqTrebleZ2L, b0_t, b1_t, b2_t, a1_t, a2_t)
        rightSample = RiffBiquadProcess(rightSample, rVoices(voiceIndex).EqTrebleZ1R, rVoices(voiceIndex).EqTrebleZ2R, b0_t, b1_t, b2_t, a1_t, a2_t)
    End If
End Sub

'/**
' * @function RiffProcessFreeverb
' * @brief Applies a Freeverb-style damped comb network with stereo cross-feed.
' */
Private Sub RiffProcessFreeverb(ByVal voiceIndex As Long, ByVal baseIndex As Long, ByVal writeIndex As Long, ByVal mix As Single, ByVal decay As Single, ByRef leftSample As Single, ByRef rightSample As Single, ByRef feedbackLeft As Single, ByRef feedbackRight As Single)
    Static lastVoice As Long
    Static lastDecay As Single
    Static lastMix As Single
    Static fb As Single
    Static damp As Single
    Static invDamp As Single
    Static clampedMix As Single
    
    Dim recalc As Boolean
    If voiceIndex <> lastVoice Then
        recalc = True
        lastVoice = voiceIndex
    End If
    
    If recalc Or decay <> lastDecay Then
        Dim clampDecay As Single
        clampDecay = decay
        If clampDecay < 0! Then clampDecay = 0!
        If clampDecay > RIFF_UNITY_GAIN Then clampDecay = RIFF_UNITY_GAIN
        
        fb = RIFF_REVERB_FEEDBACK_BASE + (clampDecay * RIFF_REVERB_FEEDBACK_RANGE)
        damp = RIFF_REVERB_DAMP_BASE + ((RIFF_UNITY_GAIN - clampDecay) * RIFF_REVERB_DAMP_RANGE)
        invDamp = 1! - damp
        lastDecay = decay
    End If
    
    If recalc Or mix <> lastMix Then
        clampedMix = mix
        If clampedMix < 0! Then clampedMix = 0!
        If clampedMix > 1! Then clampedMix = 1!
        lastMix = mix
    End If

    Dim idx1 As Long, idx2 As Long, idx3 As Long, idx4 As Long
    
    idx1 = writeIndex - rVoices(voiceIndex).RevTap1
    If idx1 < 0 Then idx1 = idx1 + RIFF_RING_SAMPLES_PER_VOICE
    
    idx2 = writeIndex - rVoices(voiceIndex).RevTap2
    If idx2 < 0 Then idx2 = idx2 + RIFF_RING_SAMPLES_PER_VOICE
    
    idx3 = writeIndex - rVoices(voiceIndex).RevTap3
    If idx3 < 0 Then idx3 = idx3 + RIFF_RING_SAMPLES_PER_VOICE
    
    idx4 = writeIndex - rVoices(voiceIndex).RevTap4
    If idx4 < 0 Then idx4 = idx4 + RIFF_RING_SAMPLES_PER_VOICE

    Dim l1 As Single, r1 As Single, l2 As Single, r2 As Single
    Dim l3 As Single, r3 As Single, l4 As Single, r4 As Single

    l1 = rRingBuf(baseIndex + idx1)
    r1 = rRingBuf(baseIndex + idx1 + 1)
    
    l2 = rRingBuf(baseIndex + idx2)
    r2 = rRingBuf(baseIndex + idx2 + 1)
    
    l3 = rRingBuf(baseIndex + idx3)
    r3 = rRingBuf(baseIndex + idx3 + 1)
    
    l4 = rRingBuf(baseIndex + idx4)
    r4 = rRingBuf(baseIndex + idx4 + 1)

    rVoices(voiceIndex).RevDamp1L = (l1 * invDamp) + (rVoices(voiceIndex).RevDamp1L * damp)
    rVoices(voiceIndex).RevDamp1R = (r1 * invDamp) + (rVoices(voiceIndex).RevDamp1R * damp)
    rVoices(voiceIndex).RevDamp2L = (l2 * invDamp) + (rVoices(voiceIndex).RevDamp2L * damp)
    rVoices(voiceIndex).RevDamp2R = (r2 * invDamp) + (rVoices(voiceIndex).RevDamp2R * damp)
    rVoices(voiceIndex).RevDamp3L = (l3 * invDamp) + (rVoices(voiceIndex).RevDamp3L * damp)
    rVoices(voiceIndex).RevDamp3R = (r3 * invDamp) + (rVoices(voiceIndex).RevDamp3R * damp)
    rVoices(voiceIndex).RevDamp4L = (l4 * invDamp) + (rVoices(voiceIndex).RevDamp4L * damp)
    rVoices(voiceIndex).RevDamp4R = (r4 * invDamp) + (rVoices(voiceIndex).RevDamp4R * damp)

    Dim wetL As Single, wetR As Single
    wetL = ((rVoices(voiceIndex).RevDamp1L + rVoices(voiceIndex).RevDamp2L + rVoices(voiceIndex).RevDamp3L + rVoices(voiceIndex).RevDamp4L) * RIFF_REVERB_DIRECT_WET) + ((rVoices(voiceIndex).RevDamp2R + rVoices(voiceIndex).RevDamp4R) * RIFF_REVERB_CROSS_WET)
    wetR = ((rVoices(voiceIndex).RevDamp1R + rVoices(voiceIndex).RevDamp2R + rVoices(voiceIndex).RevDamp3R + rVoices(voiceIndex).RevDamp4R) * RIFF_REVERB_DIRECT_WET) + ((rVoices(voiceIndex).RevDamp1L + rVoices(voiceIndex).RevDamp3L) * RIFF_REVERB_CROSS_WET)

    leftSample = leftSample + (wetL * clampedMix)
    rightSample = rightSample + (wetR * clampedMix)
    feedbackLeft = feedbackLeft + (wetL * fb)
    feedbackRight = feedbackRight + (wetR * fb)
End Sub

'/**
' * @function RiffEnsureRenderTimer
' * @brief Starts the render timer only when it is not already running.
' * @return {Boolean} True if the timer is available.
' */
Private Function RiffEnsureRenderTimer() As Boolean
    If Not rCtx.Initialized Then
        RiffSetLastError RiffErrorNotInitialized
        Exit Function
    End If

    If rCtx.TimerID <> 0 Then
        RiffEnsureRenderTimer = True
        Exit Function
    End If

    If rCtx.ThunkTimerCB = 0 Then
        RiffSetLastError RiffErrorComFailure
        Exit Function
    End If

    If Not rCtx.TimerResolutionActive Then
        timeBeginPeriod RIFF_TIMER_RESOLUTION_MS
        rCtx.TimerResolutionActive = True
    End If

    rCtx.TimerID = SetTimer(0, 0, rCtx.RenderPeriodMs, rCtx.ThunkTimerCB)

    If rCtx.TimerID = 0 Then
        If rCtx.TimerResolutionActive Then
            timeEndPeriod RIFF_TIMER_RESOLUTION_MS
            rCtx.TimerResolutionActive = False
        End If
        RiffSetLastError RiffErrorComFailure
        Exit Function
    End If

    rCtx.IdleTimerTicks = 0
    RiffEnsureRenderTimer = True
End Function

'/**
' * @function RiffStopRenderTimer
' * @brief Stops the render timer and releases the high-resolution timer request without freeing audio buffers.
' */
Private Sub RiffStopRenderTimer()
    If rCtx.TimerID <> 0 Then
        KillTimer 0, rCtx.TimerID
        rCtx.TimerID = 0
    End If

    If rCtx.TimerResolutionActive Then
        timeEndPeriod RIFF_TIMER_RESOLUTION_MS
        rCtx.TimerResolutionActive = False
    End If

    rCtx.IdleTimerTicks = 0
End Sub

'/**
' * @function RiffHasActiveVoices
' * @brief Returns True when at least one voice is currently active and not fully released.
' * @return {Boolean} True if playback is still active.
' */
Private Function RiffHasActiveVoices() As Boolean
    Dim i As Long

    For i = 0 To RIFF_MAX_VOICE_INDEX
        If rVoices(i).Active And rVoices(i).Playing Then
            RiffHasActiveVoices = True
            Exit Function
        End If
    Next i
End Function

#If VBA7 Then
Private Sub RiffTimerCallback(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Private Sub RiffTimerCallback(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    If rCtx.MagicCookie <> RIFF_MAGIC_COOKIE Then Exit Sub
    If rCtx.TimerCallbackActive Then Exit Sub
    If rCtx.TimerID <> 0 Then
        If idEvent <> rCtx.TimerID Then Exit Sub
    End If

    rCtx.TimerCallbackActive = True
    
    Dim padding As Long
    Dim framesAvailable As Long
    Dim hr As Long
    Dim bytesToWrite As Long
    Dim i As Long
    Dim frame As Long
    Dim readPos As Long
    
    Dim nBlockAlign As Integer
    Dim wBits As Integer
    Dim nChannels As Integer
    
    Dim vL As Single
    Dim vR As Single
    Dim dist As Single
    Dim lp As Single
    Dim hp As Single
    Dim ptch As Double
    
    Dim sWidth As Single
    Dim trmRate As Single
    Dim trmDepth As Single
    Dim trmPhase As Double
    Dim trmStep As Double
    
    Dim cDepth As Single
    Dim cRate As Single
    Dim cPhase As Double
    Dim cStep As Double
    
    Dim rmFreq As Single
    Dim rmMix As Single
    Dim rmPhase As Double
    Dim rmStep As Double
    
    Dim apRate As Single
    Dim apDepth As Single
    Dim apPhase As Double
    Dim apStep As Double
    
    Dim flgRate As Single
    Dim flgDepth As Single
    Dim flgFB As Single
    Dim flgPhase As Double
    Dim flgStep As Double
    
    Dim eqB As Single
    Dim eqM As Single
    Dim eqT As Single
    Dim eqAlphaLow As Single
    Dim eqAlphaHigh As Single
    
    Dim cmpThresh As Single
    Dim cmpRatio As Single
    Dim cmpEnv As Single
    Dim cmpGain As Single
    
    Dim bdSteps As Single
    Dim dsFactor As Long
    Dim dsCount As Long
    Dim lastL As Single
    Dim lastR As Single
    
    Dim rMix As Single
    Dim rTime As Single
    Dim rt1 As Long
    Dim rt2 As Long
    Dim rt3 As Long
    Dim rt4 As Long
    
    Dim dTime As Single
    Dim dFB As Single
    Dim dMix As Single
    Dim dSamples As Long
    Dim dWrite As Long
    Dim dBase As Long
    
    Dim align As Long
    Dim ptchAlign As Double
    Dim loopSnd As Boolean
    Dim pos As Double
    Dim loopStart As Double
    Dim loopEnd As Double
    
    Dim fadeState As Long
    Dim fadeCur As Long
    Dim fadeTot As Long
    Dim fadeMult As Single
    
    Dim fL As Single
    Dim fR As Single
    Dim l1 As Long
    Dim l2 As Long
    
    Dim srcIdx As Double
    Dim writeIdx As Long
    
    Dim framesNeeded As Long
    Dim bytesNeeded As Long
    Dim bytesAvail As Long
    Dim remBytes As Long
    Dim sampleCount32 As Long
    Dim sampleCount16 As Long
    Dim sourceSampleCount As Long
    
    #If VBA7 Then
        Dim pData As LongPtr
        Dim ptr As LongPtr
    #Else
        Dim pData As Long
        Dim ptr As Long
    #End If

    If Not RiffEngineHasActivePlayback() Then
        rCtx.MasterPeakL = rCtx.MasterPeakL * RIFF_MASTER_IDLE_PEAK_DECAY
        rCtx.MasterPeakR = rCtx.MasterPeakR * RIFF_MASTER_IDLE_PEAK_DECAY

        If rCtx.AutoSuspendTimer Then
            rCtx.IdleTimerTicks = rCtx.IdleTimerTicks + 1
            If rCtx.IdleTimerTicks >= RIFF_IDLE_TIMER_STOP_TICKS Then
                RiffStopRenderTimer
            End If
        Else
            rCtx.IdleTimerTicks = 0
        End If

        rCtx.TimerCallbackActive = False
        Exit Sub
    End If
    
    hr = vCall(rCtx.AudioClient, VTI_AUDIO_CLIENT_GET_CURRENT_PADDING, VarPtr(padding))
    If hr <> 0 Then
        rCtx.TimerCallbackActive = False
        Exit Sub
    End If
    
    framesAvailable = rCtx.BufferSize - padding
    If framesAvailable <= 0 Then
        rCtx.TimerCallbackActive = False
        Exit Sub
    End If

    If rCtx.MaxWriteFrames > 0 Then
        If framesAvailable > rCtx.MaxWriteFrames Then
            framesAvailable = rCtx.MaxWriteFrames
        End If
    End If
    
    RiffTickBusFades framesAvailable
    
    hr = vCall(rCtx.RenderClient, VTI_AUDIO_RENDER_CLIENT_GET_BUFFER, framesAvailable, VarPtr(pData))
    If hr <> 0 Or pData = 0 Then
        rCtx.TimerCallbackActive = False
        Exit Sub
    End If
    
    RtlMoveMemory VarPtr(nChannels), ByVal (rCtx.MixFormatPtr + RIFF_WFX_CHANNELS_OFFSET), LenB(nChannels)
    RtlMoveMemory VarPtr(nBlockAlign), ByVal (rCtx.MixFormatPtr + RIFF_WFX_BLOCK_ALIGN_OFFSET), LenB(nBlockAlign)
    RtlMoveMemory VarPtr(wBits), ByVal (rCtx.MixFormatPtr + RIFF_WFX_BITS_OFFSET), LenB(wBits)
    
    bytesToWrite = framesAvailable * CLng(nBlockAlign)
    align = CLng(nBlockAlign)
    
    eqAlphaLow = 1! - Exp(-PI2 * RIFF_EQ_LOW_CROSSOVER_HZ / CSng(rCtx.sampleRate))
    eqAlphaHigh = 1! - Exp(-PI2 * RIFF_EQ_HIGH_CROSSOVER_HZ / CSng(rCtx.sampleRate))
    
    rCtx.MasterPeakL = rCtx.MasterPeakL * RIFF_ACTIVE_PEAK_DECAY
    rCtx.MasterPeakR = rCtx.MasterPeakR * RIFF_ACTIVE_PEAK_DECAY
    
    Dim currentMasterPeakL As Single
    Dim currentMasterPeakR As Single
    currentMasterPeakL = 0!
    currentMasterPeakR = 0!
    Dim isMixFloat32 As Boolean
    isMixFloat32 = RiffMixFormatIsFloat32()

    If wBits = RIFF_FLOAT32_BITS Then
        sampleCount32 = bytesToWrite \ RIFF_SINGLE_BYTES
        RiffEnsureSingleScratch rMixArr32, rMixArr32Cap, sampleCount32
        RiffClearSingleScratch rMixArr32, sampleCount32
        
        For i = 0 To RIFF_MAX_VOICE_INDEX
            If rVoices(i).Active And rVoices(i).Playing Then
                
                If rVoices(i).Paused Then
                    GoTo NextVoice32
                End If
                
                Dim isSourceValid As Boolean
                isSourceValid = False
                
                If rVoices(i).IsOscillator Then
                    isSourceValid = True
                ElseIf rVoices(i).BufferIndex >= 0 Then
                    If rCtx.Buffers(rVoices(i).BufferIndex).Active Then
                        isSourceValid = True
                    End If
                End If
                
                If isSourceValid Then
                    pos = rVoices(i).Position
                    ptch = rVoices(i).Pitch
                    loopSnd = rVoices(i).Looping
                    loopStart = rVoices(i).loopStart
                    loopEnd = rVoices(i).loopEnd

                    framesNeeded = Int(framesAvailable * ptch) + RIFF_WAV_EXPORT_CHANNELS
                    bytesNeeded = framesNeeded * align
                    sourceSampleCount = bytesNeeded \ 4
                    
                    If Not rVoices(i).IsOscillator Then
                        ptr = rCtx.Buffers(rVoices(i).BufferIndex).BufferPtr
                        readPos = (CLng(pos) \ align) * align
                        
                        If readPos < 0 Then readPos = 0
                        
                        bytesAvail = CLng(loopEnd) - readPos
                        If bytesAvail < 0 Then bytesAvail = 0
                        
                        If isMixFloat32 Then
                            RiffEnsureSingleScratch rSrcArr32, rSrcArr32Cap, sourceSampleCount
                            RiffClearSingleScratch rSrcArr32, sourceSampleCount
                        Else
                            RiffEnsureLongScratch rSrcArrI32, rSrcArrI32Cap, sourceSampleCount
                            RiffClearLongScratch rSrcArrI32, sourceSampleCount
                        End If

                        If bytesNeeded <= bytesAvail Then
                            If isMixFloat32 Then
                                RtlMoveMemoryToSingle rSrcArr32(0), ByVal (ptr + readPos), bytesNeeded
                            Else
                                RtlMoveMemory VarPtr(rSrcArrI32(0)), ByVal (ptr + readPos), bytesNeeded
                            End If
                        Else
                            If bytesAvail > 0 Then
                                If isMixFloat32 Then
                                    RtlMoveMemoryToSingle rSrcArr32(0), ByVal (ptr + readPos), bytesAvail
                                Else
                                    RtlMoveMemory VarPtr(rSrcArrI32(0)), ByVal (ptr + readPos), bytesAvail
                                End If
                            End If
                            If loopSnd Then
                                remBytes = bytesNeeded - bytesAvail
                                Dim loopBytes32 As Long
                                Dim chunkBytes32 As Long
                                loopBytes32 = CLng(loopEnd - loopStart)
                                Do While remBytes > 0 And loopBytes32 > 0
                                    chunkBytes32 = remBytes
                                    If chunkBytes32 > loopBytes32 Then chunkBytes32 = loopBytes32
                                    If isMixFloat32 Then
                                        RtlMoveMemoryToSingle rSrcArr32((bytesNeeded - remBytes) \ 4), ByVal (ptr + CLng(loopStart)), chunkBytes32
                                    Else
                                        RtlMoveMemory VarPtr(rSrcArrI32((bytesNeeded - remBytes) \ 4)), ByVal (ptr + CLng(loopStart)), chunkBytes32
                                    End If
                                    remBytes = remBytes - chunkBytes32
                                Loop
                            End If
                        End If
                    End If
                    
                    dist = rVoices(i).Distortion
                    lp = rVoices(i).lowPass
                    hp = rVoices(i).highPass
                    sWidth = rVoices(i).StereoWidth
                    eqB = rVoices(i).EqBass
                    eqM = rVoices(i).EqMid
                    eqT = rVoices(i).EqTreble
                    cmpThresh = rVoices(i).CompThreshold
                    cmpRatio = rVoices(i).CompRatio
                    cmpEnv = rVoices(i).CompEnv
                    bdSteps = rVoices(i).BitcrushSteps
                    dsFactor = rVoices(i).BitcrushDownsample
                    dsCount = rVoices(i).BitcrushDsCount
                    lastL = rVoices(i).BitcrushLastL
                    lastR = rVoices(i).BitcrushLastR
                    rmFreq = rVoices(i).RingModFreq
                    rmMix = rVoices(i).RingModMix
                    rmPhase = rVoices(i).RingModPhase
                    rmStep = (PI2 * rmFreq) / CDbl(rCtx.sampleRate)
                    trmRate = rVoices(i).TremoloRate
                    trmDepth = rVoices(i).TremoloDepth
                    trmPhase = rVoices(i).TremoloPhase
                    trmStep = (PI2 * trmRate) / CDbl(rCtx.sampleRate)
                    apRate = rVoices(i).AutoPanRate
                    apDepth = rVoices(i).AutoPanDepth
                    apPhase = rVoices(i).AutoPanPhase
                    apStep = (PI2 * apRate) / CDbl(rCtx.sampleRate)
                    cRate = rVoices(i).ChorusRate
                    cDepth = rVoices(i).ChorusDepth
                    cPhase = rVoices(i).ChorusPhase
                    cStep = (PI2 * cRate) / CDbl(rCtx.sampleRate)
                    flgRate = rVoices(i).FlangerRate
                    flgDepth = rVoices(i).FlangerDepth
                    flgFB = rVoices(i).FlangerFeedback
                    flgPhase = rVoices(i).FlangerPhase
                    flgStep = (PI2 * flgRate) / CDbl(rCtx.sampleRate)
                    rMix = rVoices(i).ReverbMix
                    rTime = rVoices(i).ReverbTime
                    rt1 = rVoices(i).RevTap1
                    rt2 = rVoices(i).RevTap2
                    rt3 = rVoices(i).RevTap3
                    rt4 = rVoices(i).RevTap4
                    dTime = rVoices(i).DelayTime
                    dFB = rVoices(i).DelayFeedback
                    dMix = rVoices(i).DelayMix
                    dSamples = Int(dTime * rCtx.sampleRate) * RIFF_WAV_EXPORT_CHANNELS
                    dWrite = rVoices(i).RingWritePos
                    dBase = i * RIFF_RING_SAMPLES_PER_VOICE
                    fadeState = rVoices(i).fadeState
                    fadeCur = rVoices(i).FadeFramesCurrent
                    fadeTot = rVoices(i).FadeFramesTotal
                    
                    If rVoices(i).IsOscillator Then
                        srcIdx = 0
                    Else
                        srcIdx = (pos - CDbl(readPos)) / CDbl(align)
                        If srcIdx < 0# Then srcIdx = 0#
                    End If
                    writeIdx = 0
                    ptchAlign = ptch * CDbl(align)
                    
                    Dim currentVoicePeakL As Single
                    Dim currentVoicePeakR As Single
                    currentVoicePeakL = 0!
                    currentVoicePeakR = 0!
                    rVoices(i).PeakL = rVoices(i).PeakL * RIFF_ACTIVE_PEAK_DECAY
                    rVoices(i).PeakR = rVoices(i).PeakR * RIFF_ACTIVE_PEAK_DECAY

                    Dim currentVoiceBus As Long
                    Dim baseVol As Single
                    currentVoiceBus = rVoices(i).busID
                    baseVol = rVoices(i).Volume * rCtx.MasterVolume * RiffBusMixVolume(currentVoiceBus)
                    
                    If nChannels = RIFF_WAV_EXPORT_CHANNELS Then
                        For frame = 0 To framesAvailable - 1
                            If Not rVoices(i).IsOscillator Then
                                If pos >= loopEnd Then
                                    If loopSnd Then
                                        pos = loopStart
                                    Else
                                        rVoices(i).Playing = False
                                        rVoices(i).Active = False
                                        Exit For
                                    End If
                                End If
                            End If
                            
                            If fadeState = 1 Then
                                If fadeCur < fadeTot Then
                                    fadeCur = fadeCur + 1
                                    fadeMult = CSng(fadeCur) / CSng(fadeTot)
                                Else
                                    fadeState = 0
                                    fadeMult = 1!
                                End If
                            ElseIf fadeState = 2 Then
                                If fadeCur < fadeTot Then
                                    fadeCur = fadeCur + 1
                                    fadeMult = 1! - (CSng(fadeCur) / CSng(fadeTot))
                                Else
                                    rVoices(i).Playing = False
                                    rVoices(i).Active = False
                                    Exit For
                                End If
                            Else
                                fadeMult = 1!
                            End If
                            
                            If rVoices(i).IsOscillator Then
                                fL = RiffNextOscillatorSample(i)
                                fR = fL
                            Else
                                Dim sBase32 As Long
                                Dim sID As Long
                                Dim sFrac32 As Single
                                sBase32 = Int(srcIdx)
                                sID = sBase32 * RIFF_WAV_EXPORT_CHANNELS
                                sFrac32 = CSng(srcIdx - CDbl(sBase32))
                                If isMixFloat32 Then
                                    If sID + 3 < sourceSampleCount Then
                                        fL = rSrcArr32(sID) + ((rSrcArr32(sID + 2) - rSrcArr32(sID)) * sFrac32)
                                        fR = rSrcArr32(sID + 1) + ((rSrcArr32(sID + 3) - rSrcArr32(sID + 1)) * sFrac32)
                                    ElseIf sID + 1 < sourceSampleCount Then
                                        fL = rSrcArr32(sID)
                                        fR = rSrcArr32(sID + 1)
                                    Else
                                        fL = 0!
                                        fR = 0!
                                    End If
                                ElseIf sID + 3 < sourceSampleCount Then
                                    fL = CSng((CDbl(rSrcArrI32(sID)) + ((CDbl(rSrcArrI32(sID + 2)) - CDbl(rSrcArrI32(sID))) * CDbl(sFrac32))) / RIFF_PCM32_SCALE)
                                    fR = CSng((CDbl(rSrcArrI32(sID + 1)) + ((CDbl(rSrcArrI32(sID + 3)) - CDbl(rSrcArrI32(sID + 1))) * CDbl(sFrac32))) / RIFF_PCM32_SCALE)
                                ElseIf sID + 1 < sourceSampleCount Then
                                    fL = CSng(CDbl(rSrcArrI32(sID)) / RIFF_PCM32_SCALE)
                                    fR = CSng(CDbl(rSrcArrI32(sID + 1)) / RIFF_PCM32_SCALE)
                                Else
                                    fL = 0!
                                    fR = 0!
                                End If
                            End If
                            
                            If dsFactor > 1 Then
                                If dsCount >= dsFactor Then
                                    dsCount = 0
                                    lastL = fL
                                    lastR = fR
                                Else
                                    fL = lastL
                                    fR = lastR
                                    dsCount = dsCount + 1
                                End If
                            End If
                            
                            If bdSteps > 0! Then
                                fL = Fix(fL * bdSteps) / bdSteps
                                fR = Fix(fR * bdSteps) / bdSteps
                            End If
                            
                            fL = fL * dist
                            If fL > 1! Then
                                fL = 1!
                            ElseIf fL < -1! Then
                                fL = -1!
                            End If
                            
                            fR = fR * dist
                            If fR > 1! Then
                                fR = 1!
                            ElseIf fR < -1! Then
                                fR = -1!
                            End If
                            
                            RiffProcessVoiceFilters i, fL, fR, lp, hp, eqB, eqM, eqT
                            
                            If rmMix > 0! Then
                                Dim rmOsc As Single
                                rmOsc = Sin(rmPhase)
                                fL = fL * (1! - rmMix) + (fL * rmOsc) * rmMix
                                fR = fR * (1! - rmMix) + (fR * rmOsc) * rmMix
                                rmPhase = rmPhase + rmStep
                                If rmPhase > PI2 Then rmPhase = rmPhase - PI2
                            End If
                            
                            If trmDepth > 0! Then
                                Dim trmMult As Single
                                trmMult = RIFF_UNITY_GAIN - trmDepth * (RIFF_HALF_SCALE + RIFF_HALF_SCALE * CSng(Sin(trmPhase)))
                                fL = fL * trmMult
                                fR = fR * trmMult
                                trmPhase = trmPhase + trmStep
                                If trmPhase > PI2 Then trmPhase = trmPhase - PI2
                            End If
                            
                            If sWidth <> 1! Then
                                Dim midS As Single
                                Dim sideS As Single
                                midS = (fL + fR) * RIFF_HALF_SCALE
                                sideS = (fL - fR) * RIFF_HALF_SCALE
                                fL = midS + sideS * sWidth
                                fR = midS - sideS * sWidth
                            End If
                            
                            Dim bufInL As Single
                            Dim bufInR As Single
                            bufInL = fL
                            bufInR = fR
                            
                            If flgDepth > 0! Then
                                Dim fDel As Long
                                Dim fRd As Long
                                Dim flgL As Single
                                Dim flgR As Single
                                
                                fDel = Int((RIFF_FLANGER_BASE_DELAY_SEC + RIFF_FLANGER_MOD_DELAY_SEC * CSng(Sin(flgPhase))) * rCtx.sampleRate) * RIFF_WAV_EXPORT_CHANNELS
                                fRd = dWrite - fDel
                                If fRd < 0 Then fRd = fRd + RIFF_RING_SAMPLES_PER_VOICE
                                
                                flgL = rRingBuf(dBase + fRd)
                                flgR = rRingBuf(dBase + fRd + 1)
                                
                                fL = fL + flgL * flgDepth
                                fR = fR + flgR * flgDepth
                                bufInL = bufInL + flgL * flgFB
                                bufInR = bufInR + flgR * flgFB
                                
                                flgPhase = flgPhase + flgStep
                                If flgPhase > PI2 Then flgPhase = flgPhase - PI2
                            End If
                            
                            If cDepth > 0! Then
                                Dim cDelay As Long
                                Dim cRead As Long
                                
                                cDelay = Int((RIFF_CHORUS_BASE_DELAY_SEC + RIFF_CHORUS_MOD_DELAY_SEC * CSng(Sin(cPhase))) * rCtx.sampleRate) * RIFF_WAV_EXPORT_CHANNELS
                                cRead = dWrite - cDelay
                                If cRead < 0 Then cRead = cRead + RIFF_RING_SAMPLES_PER_VOICE
                                
                                fL = fL * (RIFF_UNITY_GAIN - cDepth * RIFF_HALF_SCALE) + rRingBuf(dBase + cRead) * cDepth
                                fR = fR * (RIFF_UNITY_GAIN - cDepth * RIFF_HALF_SCALE) + rRingBuf(dBase + cRead + 1) * cDepth
                                cPhase = cPhase + cStep
                                If cPhase > PI2 Then cPhase = cPhase - PI2
                            End If
                            
                            If dMix > 0! And dSamples > 0 Then
                                Dim dRead As Long
                                Dim dL As Single
                                Dim dR As Single
                                
                                dRead = dWrite - dSamples
                                If dRead < 0 Then dRead = dRead + RIFF_RING_SAMPLES_PER_VOICE
                                
                                dL = rRingBuf(dBase + dRead)
                                dR = rRingBuf(dBase + dRead + 1)
                                fL = fL + dL * dMix
                                fR = fR + dR * dMix
                                bufInL = bufInL + dL * dFB
                                bufInR = bufInR + dR * dFB
                            End If
                            
                            If rMix > 0! Then
                                RiffProcessFreeverb i, dBase, dWrite, rMix, rTime, fL, fR, bufInL, bufInR
                            End If
                            
                            rRingBuf(dBase + dWrite) = bufInL
                            rRingBuf(dBase + dWrite + 1) = bufInR
                            dWrite = dWrite + RIFF_WAV_EXPORT_CHANNELS
                            If dWrite >= RIFF_RING_SAMPLES_PER_VOICE Then dWrite = 0
                            
                            If cmpRatio > 1! Then
                                Dim pkL As Single
                                Dim pkR As Single
                                Dim maxPk As Single
                                
                                pkL = Abs(fL)
                                pkR = Abs(fR)
                                maxPk = pkL
                                If pkR > maxPk Then maxPk = pkR
                                
                                If maxPk > cmpEnv Then
                                    cmpEnv = cmpEnv + RIFF_COMP_ATTACK_COEFF * (maxPk - cmpEnv)
                                Else
                                    cmpEnv = cmpEnv + RIFF_COMP_RELEASE_COEFF * (maxPk - cmpEnv)
                                End If
                                
                                If cmpEnv > cmpThresh And cmpEnv > RIFF_COMP_ENV_FLOOR Then
                                    cmpGain = cmpThresh + ((cmpEnv - cmpThresh) / cmpRatio)
                                    cmpGain = cmpGain / cmpEnv
                                    fL = fL * cmpGain
                                    fR = fR * cmpGain
                                End If
                            End If
                            
                            Dim curPan As Single
                            curPan = rVoices(i).Pan
                            If apDepth > 0! Then
                                curPan = curPan + (Sin(apPhase) * apDepth)
                                If curPan > 1! Then
                                    curPan = 1!
                                ElseIf curPan < -1! Then
                                    curPan = -1!
                                End If
                                apPhase = apPhase + apStep
                                If apPhase > PI2 Then apPhase = apPhase - PI2
                            End If
                            
                            vL = baseVol
                            vR = baseVol
                            
                            If curPan > 0! Then vL = vL * (1! - curPan)
                            If curPan < 0! Then vR = vR * (1! + curPan)
                            
                            fL = fL * vL * fadeMult
                            fR = fR * vR * fadeMult
                            
                            If Abs(fL) > currentVoicePeakL Then currentVoicePeakL = Abs(fL)
                            If Abs(fR) > currentVoicePeakR Then currentVoicePeakR = Abs(fR)
                            
                            rMixArr32(writeIdx) = rMixArr32(writeIdx) + fL
                            rMixArr32(writeIdx + 1) = rMixArr32(writeIdx + 1) + fR

                            If Abs(rMixArr32(writeIdx)) > currentMasterPeakL Then currentMasterPeakL = Abs(rMixArr32(writeIdx))
                            If Abs(rMixArr32(writeIdx + 1)) > currentMasterPeakR Then currentMasterPeakR = Abs(rMixArr32(writeIdx + 1))
                            
                            writeIdx = writeIdx + RIFF_WAV_EXPORT_CHANNELS
                            srcIdx = srcIdx + ptch
                            pos = pos + ptchAlign
                        Next frame
                    End If
                    
                    If currentVoicePeakL > rVoices(i).PeakL Then rVoices(i).PeakL = currentVoicePeakL
                    If currentVoicePeakR > rVoices(i).PeakR Then rVoices(i).PeakR = currentVoicePeakR
                    
                    If currentVoicePeakL > rCtx.Buses(currentVoiceBus).PeakL Then rCtx.Buses(currentVoiceBus).PeakL = currentVoicePeakL
                    If currentVoicePeakR > rCtx.Buses(currentVoiceBus).PeakR Then rCtx.Buses(currentVoiceBus).PeakR = currentVoicePeakR
                    
                    rVoices(i).Position = pos
                    rVoices(i).TremoloPhase = trmPhase
                    rVoices(i).AutoPanPhase = apPhase
                    rVoices(i).RingModPhase = rmPhase
                    rVoices(i).ChorusPhase = cPhase
                    rVoices(i).FlangerPhase = flgPhase
                    rVoices(i).RingWritePos = dWrite
                    rVoices(i).CompEnv = cmpEnv
                    rVoices(i).BitcrushDsCount = dsCount
                    rVoices(i).BitcrushLastL = lastL
                    rVoices(i).BitcrushLastR = lastR
                    rVoices(i).fadeState = fadeState
                    rVoices(i).FadeFramesCurrent = fadeCur
                End If
            End If
NextVoice32:
        Next i
        
        If currentMasterPeakL > rCtx.MasterPeakL Then rCtx.MasterPeakL = currentMasterPeakL
        If currentMasterPeakR > rCtx.MasterPeakR Then rCtx.MasterPeakR = currentMasterPeakR
        
        For frame = 0 To sampleCount32 - 1
            If rMixArr32(frame) > 1! Then
                rMixArr32(frame) = 1!
            ElseIf rMixArr32(frame) < -1! Then
                rMixArr32(frame) = -1!
            End If
        Next frame

        If isMixFloat32 Then
            RtlMoveMemory ByVal pData, VarPtr(rMixArr32(0)), bytesToWrite
        Else
            RiffEnsureLongScratch rMixInt32, rMixInt32Cap, sampleCount32
            For frame = 0 To sampleCount32 - 1
                If rMixArr32(frame) >= 1! Then
                    rMixInt32(frame) = CLng(RIFF_PCM32_MAX)
                ElseIf rMixArr32(frame) <= -1! Then
                    rMixInt32(frame) = -CLng(RIFF_PCM32_MAX)
                Else
                    rMixInt32(frame) = CLng(rMixArr32(frame) * RIFF_PCM32_MAX)
                End If
            Next frame
            RtlMoveMemory ByVal pData, VarPtr(rMixInt32(0)), bytesToWrite
        End If

    ElseIf wBits = RIFF_PCM16_BITS Then
        sampleCount16 = bytesToWrite \ RIFF_INTEGER_BYTES
        RiffEnsureIntegerScratch rMixArr16, rMixArr16Cap, sampleCount16
        RiffClearIntegerScratch rMixArr16, sampleCount16
        
        For i = 0 To RIFF_MAX_VOICE_INDEX
            If rVoices(i).Active And rVoices(i).Playing Then
                
                If rVoices(i).Paused Then
                    GoTo NextVoice16
                End If
                
                Dim isSourceValid16 As Boolean
                isSourceValid16 = False
                
                If rVoices(i).IsOscillator Then
                    isSourceValid16 = True
                ElseIf rVoices(i).BufferIndex >= 0 Then
                    If rCtx.Buffers(rVoices(i).BufferIndex).Active Then
                        isSourceValid16 = True
                    End If
                End If
                
                If isSourceValid16 Then
                    pos = rVoices(i).Position
                    ptch = rVoices(i).Pitch
                    loopSnd = rVoices(i).Looping
                    loopStart = rVoices(i).loopStart
                    loopEnd = rVoices(i).loopEnd
                    
                    framesNeeded = Int(framesAvailable * ptch) + RIFF_WAV_EXPORT_CHANNELS
                    bytesNeeded = framesNeeded * align
                    sourceSampleCount = bytesNeeded \ 2
                    
                    If Not rVoices(i).IsOscillator Then
                        ptr = rCtx.Buffers(rVoices(i).BufferIndex).BufferPtr
                        readPos = (CLng(pos) \ align) * align
                        
                        If readPos < 0 Then readPos = 0
                        
                        bytesAvail = CLng(loopEnd) - readPos
                        If bytesAvail < 0 Then bytesAvail = 0
                        
                        RiffEnsureIntegerScratch rSrcArr16, rSrcArr16Cap, sourceSampleCount
                        RiffClearIntegerScratch rSrcArr16, sourceSampleCount

                        If bytesNeeded <= bytesAvail Then
                            RtlMoveMemoryToInteger rSrcArr16(0), ByVal (ptr + readPos), bytesNeeded
                        Else
                            If bytesAvail > 0 Then
                                RtlMoveMemoryToInteger rSrcArr16(0), ByVal (ptr + readPos), bytesAvail
                            End If
                            If loopSnd Then
                                remBytes = bytesNeeded - bytesAvail
                                Dim loopBytes16 As Long
                                Dim chunkBytes16 As Long
                                loopBytes16 = CLng(loopEnd - loopStart)
                                Do While remBytes > 0 And loopBytes16 > 0
                                    chunkBytes16 = remBytes
                                    If chunkBytes16 > loopBytes16 Then chunkBytes16 = loopBytes16
                                    RtlMoveMemoryToInteger rSrcArr16((bytesNeeded - remBytes) \ 2), ByVal (ptr + CLng(loopStart)), chunkBytes16
                                    remBytes = remBytes - chunkBytes16
                                Loop
                            End If
                        End If
                    End If
                    
                    dist = rVoices(i).Distortion
                    lp = rVoices(i).lowPass
                    hp = rVoices(i).highPass
                    sWidth = rVoices(i).StereoWidth
                    eqB = rVoices(i).EqBass
                    eqM = rVoices(i).EqMid
                    eqT = rVoices(i).EqTreble
                    cmpThresh = rVoices(i).CompThreshold
                    cmpRatio = rVoices(i).CompRatio
                    cmpEnv = rVoices(i).CompEnv
                    bdSteps = rVoices(i).BitcrushSteps
                    dsFactor = rVoices(i).BitcrushDownsample
                    dsCount = rVoices(i).BitcrushDsCount
                    lastL = rVoices(i).BitcrushLastL
                    lastR = rVoices(i).BitcrushLastR
                    rmFreq = rVoices(i).RingModFreq
                    rmMix = rVoices(i).RingModMix
                    rmPhase = rVoices(i).RingModPhase
                    rmStep = (PI2 * rmFreq) / CDbl(rCtx.sampleRate)
                    trmRate = rVoices(i).TremoloRate
                    trmDepth = rVoices(i).TremoloDepth
                    trmPhase = rVoices(i).TremoloPhase
                    trmStep = (PI2 * trmRate) / CDbl(rCtx.sampleRate)
                    apRate = rVoices(i).AutoPanRate
                    apDepth = rVoices(i).AutoPanDepth
                    apPhase = rVoices(i).AutoPanPhase
                    apStep = (PI2 * apRate) / CDbl(rCtx.sampleRate)
                    cRate = rVoices(i).ChorusRate
                    cDepth = rVoices(i).ChorusDepth
                    cPhase = rVoices(i).ChorusPhase
                    cStep = (PI2 * cRate) / CDbl(rCtx.sampleRate)
                    flgRate = rVoices(i).FlangerRate
                    flgDepth = rVoices(i).FlangerDepth
                    flgFB = rVoices(i).FlangerFeedback
                    flgPhase = rVoices(i).FlangerPhase
                    flgStep = (PI2 * flgRate) / CDbl(rCtx.sampleRate)
                    rMix = rVoices(i).ReverbMix
                    rTime = rVoices(i).ReverbTime
                    rt1 = rVoices(i).RevTap1
                    rt2 = rVoices(i).RevTap2
                    rt3 = rVoices(i).RevTap3
                    rt4 = rVoices(i).RevTap4
                    dTime = rVoices(i).DelayTime
                    dFB = rVoices(i).DelayFeedback
                    dMix = rVoices(i).DelayMix
                    dSamples = Int(dTime * rCtx.sampleRate) * RIFF_WAV_EXPORT_CHANNELS
                    dWrite = rVoices(i).RingWritePos
                    dBase = i * RIFF_RING_SAMPLES_PER_VOICE
                    fadeState = rVoices(i).fadeState
                    fadeCur = rVoices(i).FadeFramesCurrent
                    fadeTot = rVoices(i).FadeFramesTotal
                    
                    If rVoices(i).IsOscillator Then
                        srcIdx = 0
                    Else
                        srcIdx = (pos - CDbl(readPos)) / CDbl(align)
                        If srcIdx < 0# Then srcIdx = 0#
                    End If
                    writeIdx = 0
                    ptchAlign = ptch * CDbl(align)
                    
                    Dim cVoicePeakL16 As Single
                    Dim cVoicePeakR16 As Single
                    cVoicePeakL16 = 0!
                    cVoicePeakR16 = 0!
                    rVoices(i).PeakL = rVoices(i).PeakL * RIFF_ACTIVE_PEAK_DECAY
                    rVoices(i).PeakR = rVoices(i).PeakR * RIFF_ACTIVE_PEAK_DECAY
                    
                    Dim cVoiceBus16 As Long
                    Dim bVol16 As Single
                    cVoiceBus16 = rVoices(i).busID
                    bVol16 = rVoices(i).Volume * rCtx.MasterVolume * RiffBusMixVolume(cVoiceBus16)

                    If nChannels = RIFF_WAV_EXPORT_CHANNELS Then
                        For frame = 0 To framesAvailable - 1
                            If Not rVoices(i).IsOscillator Then
                                If pos >= loopEnd Then
                                    If loopSnd Then
                                        pos = loopStart
                                    Else
                                        rVoices(i).Playing = False
                                        rVoices(i).Active = False
                                        Exit For
                                    End If
                                End If
                            End If
                            
                            If fadeState = 1 Then
                                If fadeCur < fadeTot Then
                                    fadeCur = fadeCur + 1
                                    fadeMult = CSng(fadeCur) / CSng(fadeTot)
                                Else
                                    fadeState = 0
                                    fadeMult = 1!
                                End If
                            ElseIf fadeState = 2 Then
                                If fadeCur < fadeTot Then
                                    fadeCur = fadeCur + 1
                                    fadeMult = 1! - (CSng(fadeCur) / CSng(fadeTot))
                                Else
                                    rVoices(i).Playing = False
                                    rVoices(i).Active = False
                                    Exit For
                                End If
                            Else
                                fadeMult = 1!
                            End If
                            
                            If rVoices(i).IsOscillator Then
                                fL = RiffNextOscillatorSample(i)
                                fR = fL
                            Else
                                Dim sBase16 As Long
                                Dim sID16 As Long
                                Dim sFrac16 As Single
                                sBase16 = Int(srcIdx)
                                sID16 = sBase16 * RIFF_WAV_EXPORT_CHANNELS
                                sFrac16 = CSng(srcIdx - CDbl(sBase16))
                                If sID16 + 3 < sourceSampleCount Then
                                    fL = (CSng(rSrcArr16(sID16)) + ((CSng(rSrcArr16(sID16 + 2)) - CSng(rSrcArr16(sID16))) * sFrac16)) * RIFF_PCM16_TO_FLOAT_SCALE
                                    fR = (CSng(rSrcArr16(sID16 + 1)) + ((CSng(rSrcArr16(sID16 + 3)) - CSng(rSrcArr16(sID16 + 1))) * sFrac16)) * RIFF_PCM16_TO_FLOAT_SCALE
                                ElseIf sID16 + 1 < sourceSampleCount Then
                                    fL = CSng(rSrcArr16(sID16)) * RIFF_PCM16_TO_FLOAT_SCALE
                                    fR = CSng(rSrcArr16(sID16 + 1)) * RIFF_PCM16_TO_FLOAT_SCALE
                                Else
                                    fL = 0!
                                    fR = 0!
                                End If
                            End If
                            
                            If dsFactor > 1 Then
                                If dsCount >= dsFactor Then
                                    dsCount = 0
                                    lastL = fL
                                    lastR = fR
                                Else
                                    fL = lastL
                                    fR = lastR
                                    dsCount = dsCount + 1
                                End If
                            End If
                            
                            If bdSteps > 0! Then
                                fL = Fix(fL * bdSteps) / bdSteps
                                fR = Fix(fR * bdSteps) / bdSteps
                            End If
                            
                            fL = fL * dist
                            If fL > 1! Then
                                fL = 1!
                            ElseIf fL < -1! Then
                                fL = -1!
                            End If
                            
                            fR = fR * dist
                            If fR > 1! Then
                                fR = 1!
                            ElseIf fR < -1! Then
                                fR = -1!
                            End If
                            
                            RiffProcessVoiceFilters i, fL, fR, lp, hp, eqB, eqM, eqT
                            
                            If rmMix > 0! Then
                                Dim rmOsc16 As Single
                                rmOsc16 = Sin(rmPhase)
                                fL = fL * (1! - rmMix) + (fL * rmOsc16) * rmMix
                                fR = fR * (1! - rmMix) + (fR * rmOsc16) * rmMix
                                rmPhase = rmPhase + rmStep
                                If rmPhase > PI2 Then rmPhase = rmPhase - PI2
                            End If
                            
                            If trmDepth > 0! Then
                                Dim trmM16 As Single
                                trmM16 = RIFF_UNITY_GAIN - trmDepth * (RIFF_HALF_SCALE + RIFF_HALF_SCALE * CSng(Sin(trmPhase)))
                                fL = fL * trmM16
                                fR = fR * trmM16
                                trmPhase = trmPhase + trmStep
                                If trmPhase > PI2 Then trmPhase = trmPhase - PI2
                            End If
                            
                            If sWidth <> 1! Then
                                Dim m16 As Single
                                Dim sd16 As Single
                                m16 = (fL + fR) * RIFF_HALF_SCALE
                                sd16 = (fL - fR) * RIFF_HALF_SCALE
                                fL = m16 + sd16 * sWidth
                                fR = m16 - sd16 * sWidth
                            End If
                            
                            Dim bufInL16 As Single
                            Dim bufInR16 As Single
                            bufInL16 = fL
                            bufInR16 = fR
                            
                            If flgDepth > 0! Then
                                Dim fDel16 As Long
                                Dim fRd16 As Long
                                Dim flgL16 As Single
                                Dim flgR16 As Single
                                
                                fDel16 = Int((RIFF_FLANGER_BASE_DELAY_SEC + RIFF_FLANGER_MOD_DELAY_SEC * CSng(Sin(flgPhase))) * rCtx.sampleRate) * RIFF_WAV_EXPORT_CHANNELS
                                fRd16 = dWrite - fDel16
                                If fRd16 < 0 Then fRd16 = fRd16 + RIFF_RING_SAMPLES_PER_VOICE
                                
                                flgL16 = rRingBuf(dBase + fRd16)
                                flgR16 = rRingBuf(dBase + fRd16 + 1)
                                
                                fL = fL + flgL16 * flgDepth
                                fR = fR + flgR16 * flgDepth
                                bufInL16 = bufInL16 + flgL16 * flgFB
                                bufInR16 = bufInR16 + flgR16 * flgFB
                                
                                flgPhase = flgPhase + flgStep
                                If flgPhase > PI2 Then flgPhase = flgPhase - PI2
                            End If
                            
                            If cDepth > 0! Then
                                Dim cDel16 As Long
                                Dim cRd16 As Long
                                
                                cDel16 = Int((RIFF_CHORUS_BASE_DELAY_SEC + RIFF_CHORUS_MOD_DELAY_SEC * CSng(Sin(cPhase))) * rCtx.sampleRate) * RIFF_WAV_EXPORT_CHANNELS
                                cRd16 = dWrite - cDel16
                                If cRd16 < 0 Then cRd16 = cRd16 + RIFF_RING_SAMPLES_PER_VOICE
                                
                                fL = fL * (RIFF_UNITY_GAIN - cDepth * RIFF_HALF_SCALE) + rRingBuf(dBase + cRd16) * cDepth
                                fR = fR * (RIFF_UNITY_GAIN - cDepth * RIFF_HALF_SCALE) + rRingBuf(dBase + cRd16 + 1) * cDepth
                                cPhase = cPhase + cStep
                                If cPhase > PI2 Then cPhase = cPhase - PI2
                            End If
                            
                            If dMix > 0! And dSamples > 0 Then
                                Dim dR16 As Long
                                Dim dL16 As Single
                                Dim dR16_2 As Single
                                
                                dR16 = dWrite - dSamples
                                If dR16 < 0 Then dR16 = dR16 + RIFF_RING_SAMPLES_PER_VOICE
                                
                                dL16 = rRingBuf(dBase + dR16)
                                dR16_2 = rRingBuf(dBase + dR16 + 1)
                                fL = fL + dL16 * dMix
                                fR = fR + dR16_2 * dMix
                                bufInL16 = bufInL16 + dL16 * dFB
                                bufInR16 = bufInR16 + dR16_2 * dFB
                            End If
                            
                            If rMix > 0! Then
                                RiffProcessFreeverb i, dBase, dWrite, rMix, rTime, fL, fR, bufInL16, bufInR16
                            End If
                            
                            rRingBuf(dBase + dWrite) = bufInL16
                            rRingBuf(dBase + dWrite + 1) = bufInR16
                            dWrite = dWrite + RIFF_WAV_EXPORT_CHANNELS
                            If dWrite >= RIFF_RING_SAMPLES_PER_VOICE Then dWrite = 0
                            
                            If cmpRatio > 1! Then
                                Dim pkL16 As Single
                                Dim pkR16 As Single
                                Dim maxPk16 As Single
                                
                                pkL16 = Abs(fL)
                                pkR16 = Abs(fR)
                                maxPk16 = pkL16
                                If pkR16 > maxPk16 Then maxPk16 = pkR16
                                
                                If maxPk16 > cmpEnv Then
                                    cmpEnv = cmpEnv + RIFF_COMP_ATTACK_COEFF * (maxPk16 - cmpEnv)
                                Else
                                    cmpEnv = cmpEnv + RIFF_COMP_RELEASE_COEFF * (maxPk16 - cmpEnv)
                                End If
                                
                                If cmpEnv > cmpThresh And cmpEnv > RIFF_COMP_ENV_FLOOR Then
                                    cmpGain = cmpThresh + ((cmpEnv - cmpThresh) / cmpRatio)
                                    cmpGain = cmpGain / cmpEnv
                                    fL = fL * cmpGain
                                    fR = fR * cmpGain
                                End If
                            End If
                            
                            Dim cPan16 As Single
                            cPan16 = rVoices(i).Pan
                            If apDepth > 0! Then
                                cPan16 = cPan16 + (Sin(apPhase) * apDepth)
                                If cPan16 > 1! Then
                                    cPan16 = 1!
                                ElseIf cPan16 < -1! Then
                                    cPan16 = -1!
                                End If
                                apPhase = apPhase + apStep
                                If apPhase > PI2 Then apPhase = apPhase - PI2
                            End If
                            
                            vL = bVol16
                            vR = bVol16
                            
                            If cPan16 > 0! Then vL = vL * (1! - cPan16)
                            If cPan16 < 0! Then vR = vR * (1! + cPan16)
                            
                            fL = fL * vL * fadeMult
                            fR = fR * vR * fadeMult
                            
                            If Abs(fL) > cVoicePeakL16 Then cVoicePeakL16 = Abs(fL)
                            If Abs(fR) > cVoicePeakR16 Then cVoicePeakR16 = Abs(fR)
                            
                            l1 = CLng(rMixArr16(writeIdx)) + CLng(fL * CSng(RIFF_PCM16_MAX))
                            l2 = CLng(rMixArr16(writeIdx + 1)) + CLng(fR * CSng(RIFF_PCM16_MAX))
                            
                            If l1 > RIFF_PCM16_MAX Then
                                l1 = RIFF_PCM16_MAX
                            ElseIf l1 < RIFF_PCM16_MIN Then
                                l1 = RIFF_PCM16_MIN
                            End If
                            
                            If l2 > RIFF_PCM16_MAX Then
                                l2 = RIFF_PCM16_MAX
                            ElseIf l2 < RIFF_PCM16_MIN Then
                                l2 = RIFF_PCM16_MIN
                            End If
                            
                            rMixArr16(writeIdx) = CInt(l1)
                            rMixArr16(writeIdx + 1) = CInt(l2)
                            
                            If Abs(fL) > currentMasterPeakL Then currentMasterPeakL = Abs(fL)
                            If Abs(fR) > currentMasterPeakR Then currentMasterPeakR = Abs(fR)
                            
                            writeIdx = writeIdx + RIFF_WAV_EXPORT_CHANNELS
                            srcIdx = srcIdx + ptch
                            pos = pos + ptchAlign
                        Next frame
                    End If
                    
                    If cVoicePeakL16 > rVoices(i).PeakL Then rVoices(i).PeakL = cVoicePeakL16
                    If cVoicePeakR16 > rVoices(i).PeakR Then rVoices(i).PeakR = cVoicePeakR16
                    
                    If cVoicePeakL16 > rCtx.Buses(cVoiceBus16).PeakL Then rCtx.Buses(cVoiceBus16).PeakL = cVoicePeakL16
                    If cVoicePeakR16 > rCtx.Buses(cVoiceBus16).PeakR Then rCtx.Buses(cVoiceBus16).PeakR = cVoicePeakR16
                    
                    rVoices(i).Position = pos
                    rVoices(i).TremoloPhase = trmPhase
                    rVoices(i).AutoPanPhase = apPhase
                    rVoices(i).RingModPhase = rmPhase
                    rVoices(i).ChorusPhase = cPhase
                    rVoices(i).FlangerPhase = flgPhase
                    rVoices(i).RingWritePos = dWrite
                    rVoices(i).CompEnv = cmpEnv
                    rVoices(i).BitcrushDsCount = dsCount
                    rVoices(i).BitcrushLastL = lastL
                    rVoices(i).BitcrushLastR = lastR
                    rVoices(i).fadeState = fadeState
                    rVoices(i).FadeFramesCurrent = fadeCur
                End If
            End If
NextVoice16:
        Next i
        
        If currentMasterPeakL > rCtx.MasterPeakL Then rCtx.MasterPeakL = currentMasterPeakL
        If currentMasterPeakR > rCtx.MasterPeakR Then rCtx.MasterPeakR = currentMasterPeakR
        
        RtlMoveMemory ByVal pData, VarPtr(rMixArr16(0)), bytesToWrite
    Else
        RtlZeroMemory pData, bytesToWrite
    End If
    
    vCall rCtx.RenderClient, VTI_AUDIO_RENDER_CLIENT_RELEASE_BUFFER, framesAvailable, 0&

    If rCtx.AutoSuspendTimer Then
        If RiffEngineHasActivePlayback() Then
            rCtx.IdleTimerTicks = 0
        Else
            rCtx.IdleTimerTicks = rCtx.IdleTimerTicks + 1
            If rCtx.IdleTimerTicks >= RIFF_IDLE_TIMER_STOP_TICKS Then
                RiffStopRenderTimer
            End If
        End If
    Else
        rCtx.IdleTimerTicks = 0
    End If

    rCtx.TimerCallbackActive = False
End Sub

'/**
' * @function DummyEbMode
' * @brief Safe fallback function if EbMode is inaccessible.
' * @return {Long} Always returns 1.
' */
Private Function DummyEbMode() As Long
    DummyEbMode = 1
End Function

'/**
' * @function GetAddressOf
' * @brief Extracts the memory address pointer of a VBA AddressOf callback safely.
' * @param ptr Function pointer.
' * @return {Long/LongPtr} The raw address.
' */
#If VBA7 Then
Private Function GetAddressOf(ByVal ptr As LongPtr) As LongPtr
    GetAddressOf = ptr
End Function
#Else
Private Function GetAddressOf(ByVal ptr As Long) As Long
    GetAddressOf = ptr
End Function
#End If

'/**
' * @function InitThunks
' * @brief Compiles machine code at runtime to bridge Win32 Timers and VBA logic.
' * @return {Boolean} True if successful.
' */
Private Function InitThunks() As Boolean
    #If VBA7 Then
        Const THUNK_SIZE As Long = RIFF_THUNK64_SIZE
    #Else
        Const THUNK_SIZE As Long = RIFF_THUNK32_SIZE
    #End If
    
    rCtx.ThunkTimerCB = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If rCtx.ThunkTimerCB = 0 Then
        Exit Function
    End If
    
    Dim opcodes() As Byte
    Dim hexStr As String
    Dim i As Long
    
    #If VBA7 Then
        Dim hVbe As LongPtr
        Dim pEbMode As LongPtr
        Dim pCallback As LongPtr
        Dim pKill As LongPtr
    #Else
        Dim hVbe As Long
        Dim pEbMode As Long
        Dim pCallback As Long
        Dim pKill As Long
    #End If
    
    hVbe = GetModuleHandleA("vbe7.dll")
    If hVbe = 0 Then
        hVbe = GetModuleHandleA("vba6.dll")
    End If
    
    If hVbe <> 0 Then
        pEbMode = GetProcAddress(hVbe, "EbMode")
        If pEbMode = 0 Then
            pEbMode = GetProcAddressOrdinal(hVbe, RIFF_VBE_EBMODE_ORDINAL)
        End If
    End If
    
    If pEbMode = 0 Then
        pEbMode = GetAddressOf(AddressOf DummyEbMode)
    End If
    
    pCallback = GetAddressOf(AddressOf RiffTimerCallback)
    pKill = GetProcAddress(GetModuleHandleA("user32.dll"), "KillTimer")
    
    #If Win64 Then
        hexStr = "4883EC2848894C243048895424384C894424404C894C244848B80000000000000000" & _
                 "4885C07429FFD083F801742283F802741D488B4C2430488B54244048B800000000" & _
                 "000000004885C07429FFD0EB25488B4C2430488B5424384C8B4424404C8B4C2448" & _
                 "48B800000000000000004885C07402FFD04883C428C3"
        ReDim opcodes(0 To (Len(hexStr) \ RIFF_HEX_BYTE_CHARS) - 1)
        For i = 0 To UBound(opcodes)
            opcodes(i) = CByte("&H" & Mid$(hexStr, (i * RIFF_HEX_BYTE_CHARS) + 1, RIFF_HEX_BYTE_CHARS))
        Next i
        RtlMoveMemory VarPtr(opcodes(RIFF_THUNK64_EBMODE_OFFSET)), VarPtr(pEbMode), LenB(pEbMode)
        RtlMoveMemory VarPtr(opcodes(RIFF_THUNK64_KILLTIMER_OFFSET)), VarPtr(pKill), LenB(pKill)
        RtlMoveMemory VarPtr(opcodes(RIFF_THUNK64_CALLBACK_OFFSET)), VarPtr(pCallback), LenB(pCallback)
    #Else
        hexStr = "5589E5B80000000085C07421FFD083F801741A83F80274158B4510508B450850B8" & _
                 "0000000085C0741FFFD0EB1B8B4514508B4510508B450C508B450850B800000000" & _
                 "85C07402FFD05DC21000"
        ReDim opcodes(0 To (Len(hexStr) \ RIFF_HEX_BYTE_CHARS) - 1)
        For i = 0 To UBound(opcodes)
            opcodes(i) = CByte("&H" & Mid$(hexStr, (i * RIFF_HEX_BYTE_CHARS) + 1, RIFF_HEX_BYTE_CHARS))
        Next i
        RtlMoveMemory VarPtr(opcodes(RIFF_THUNK32_EBMODE_OFFSET)), VarPtr(pEbMode), LenB(pEbMode)
        RtlMoveMemory VarPtr(opcodes(RIFF_THUNK32_KILLTIMER_OFFSET)), VarPtr(pKill), LenB(pKill)
        RtlMoveMemory VarPtr(opcodes(RIFF_THUNK32_CALLBACK_OFFSET)), VarPtr(pCallback), LenB(pCallback)
    #End If
    
    RtlMoveMemory ByVal rCtx.ThunkTimerCB, VarPtr(opcodes(0)), UBound(opcodes) + 1
    InitThunks = True
End Function

'/**
' * @function FreeThunks
' * @brief Cleans up the runtime-compiled machine code memory.
' */
Private Sub FreeThunks()
    If rCtx.ThunkTimerCB <> 0 Then
        VirtualFree rCtx.ThunkTimerCB, 0, MEM_RELEASE
    End If
    rCtx.ThunkTimerCB = 0
End Sub


'/**
' * @function RiffMixFormatIsFloat32
' * @brief Detects whether the current render format stores samples as 32-bit IEEE float.
' * @return {Boolean} True when the current WASAPI mix format is float32.
' */
Private Function RiffMixFormatIsFloat32() As Boolean
    If rCtx.MixFormatPtr = 0 Then
        Exit Function
    End If

    Dim formatTag As Integer
    Dim bitsPerSample As Integer
    Dim subFormatData1 As Long

    RtlMoveMemory VarPtr(formatTag), ByVal rCtx.MixFormatPtr, LenB(formatTag)
    RtlMoveMemory VarPtr(bitsPerSample), ByVal (rCtx.MixFormatPtr + RIFF_WFX_BITS_OFFSET), LenB(bitsPerSample)

    If bitsPerSample <> RIFF_FLOAT32_BITS Then
        Exit Function
    End If

    If formatTag = RIFF_WAVE_FORMAT_IEEE_FLOAT Then
        RiffMixFormatIsFloat32 = True
        Exit Function
    End If

    If formatTag = RIFF_WAVE_FORMAT_EXTENSIBLE Then
        RtlMoveMemory VarPtr(subFormatData1), ByVal (rCtx.MixFormatPtr + RIFF_WFX_SUBFORMAT_OFFSET), LenB(subFormatData1)
        RiffMixFormatIsFloat32 = (subFormatData1 = RIFF_WAVE_FORMAT_IEEE_FLOAT)
    End If
End Function

'/**
' * @function RiffIsSharedMixFormatSupported
' * @brief Checks a candidate shared-mode WASAPI format without initializing the audio client.
' */
#If VBA7 Then
Private Function RiffIsSharedMixFormatSupported(ByVal pFormat As LongPtr) As Boolean
    Dim pClosest As LongPtr
#Else
Private Function RiffIsSharedMixFormatSupported(ByVal pFormat As Long) As Boolean
    Dim pClosest As Long
#End If
    If rCtx.AudioClient = 0 Or pFormat = 0 Then
        Exit Function
    End If

    Dim hr As Long
    hr = vCall(rCtx.AudioClient, VTI_AUDIO_CLIENT_IS_FORMAT_SUPPORTED, AUDCLNT_SHAREMODE_SHARED, pFormat, VarPtr(pClosest))

    If pClosest <> 0 Then
        CoTaskMemFree pClosest
    End If

    RiffIsSharedMixFormatSupported = (hr = 0)
End Function

'/**
' * @function RiffTryPromoteMixFormatToFloat32
' * @brief Rewrites a WAVEFORMATEX/WAVEFORMATEXTENSIBLE structure to stereo 32-bit float while preserving sample rate.
' */
Private Sub RiffTryPromoteMixFormatToFloat32()
    If rCtx.MixFormatPtr = 0 Then
        Exit Sub
    End If

    Dim nChannels As Integer
    Dim sampleRate As Long
    Dim blockAlign As Integer
    Dim avgBytes As Long
    Dim bits As Integer
    Dim cbSize As Integer
    Dim validBits As Integer
    Dim channelMask As Long
    Dim formatTag As Integer

    RtlMoveMemory VarPtr(nChannels), ByVal (rCtx.MixFormatPtr + RIFF_WFX_CHANNELS_OFFSET), LenB(nChannels)
    RtlMoveMemory VarPtr(sampleRate), ByVal (rCtx.MixFormatPtr + RIFF_WFX_SAMPLE_RATE_OFFSET), LenB(sampleRate)
    RtlMoveMemory VarPtr(cbSize), ByVal (rCtx.MixFormatPtr + RIFF_WFX_CB_SIZE_OFFSET), LenB(cbSize)

    nChannels = RIFF_WAV_EXPORT_CHANNELS
    If sampleRate <= 0 Then
        sampleRate = RIFF_DEFAULT_SAMPLE_RATE
    End If

    bits = RIFF_FLOAT32_BITS
    blockAlign = CInt(nChannels * (bits \ 8))
    avgBytes = sampleRate * CLng(blockAlign)

    If cbSize >= RIFF_WFX_EXTENSIBLE_CB_SIZE Then
        RtlMoveMemory VarPtr(channelMask), ByVal (rCtx.MixFormatPtr + RIFF_WFX_CHANNEL_MASK_OFFSET), LenB(channelMask)
        formatTag = RIFF_WAVE_FORMAT_EXTENSIBLE
        validBits = bits
        channelMask = RIFF_STEREO_CHANNEL_MASK

        RtlMoveMemory ByVal rCtx.MixFormatPtr, VarPtr(formatTag), LenB(formatTag)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_CHANNELS_OFFSET), VarPtr(nChannels), LenB(nChannels)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_SAMPLE_RATE_OFFSET), VarPtr(sampleRate), LenB(sampleRate)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_AVG_BYTES_OFFSET), VarPtr(avgBytes), LenB(avgBytes)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_BLOCK_ALIGN_OFFSET), VarPtr(blockAlign), LenB(blockAlign)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_BITS_OFFSET), VarPtr(bits), LenB(bits)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_CB_SIZE_OFFSET), VarPtr(cbSize), LenB(cbSize)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_VALID_BITS_OFFSET), VarPtr(validBits), LenB(validBits)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_CHANNEL_MASK_OFFSET), VarPtr(channelMask), LenB(channelMask)

        Dim sf1 As Long
        Dim sf2 As Integer
        Dim sf3 As Integer
        Dim sf4(0 To RIFF_GUID_DATA4_BYTES - 1) As Byte

        sf1 = RIFF_WAVE_FORMAT_IEEE_FLOAT
        sf2 = RIFF_IEEE_FLOAT_SUBFORMAT_DATA2
        sf3 = RIFF_IEEE_FLOAT_SUBFORMAT_DATA3
        sf4(0) = RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_0
        sf4(1) = RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_1
        sf4(2) = RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_2
        sf4(3) = RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_3
        sf4(4) = RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_4
        sf4(5) = RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_5
        sf4(6) = RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_6
        sf4(7) = RIFF_IEEE_FLOAT_SUBFORMAT_DATA4_7

        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_SUBFORMAT_OFFSET), VarPtr(sf1), LenB(sf1)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_SUBFORMAT_DATA2_OFFSET), VarPtr(sf2), LenB(sf2)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_SUBFORMAT_DATA3_OFFSET), VarPtr(sf3), LenB(sf3)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_SUBFORMAT_DATA4_OFFSET), VarPtr(sf4(0)), RIFF_GUID_DATA4_BYTES
    Else
        formatTag = RIFF_WAVE_FORMAT_IEEE_FLOAT
        cbSize = 0
        RtlMoveMemory ByVal rCtx.MixFormatPtr, VarPtr(formatTag), LenB(formatTag)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_CHANNELS_OFFSET), VarPtr(nChannels), LenB(nChannels)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_SAMPLE_RATE_OFFSET), VarPtr(sampleRate), LenB(sampleRate)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_AVG_BYTES_OFFSET), VarPtr(avgBytes), LenB(avgBytes)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_BLOCK_ALIGN_OFFSET), VarPtr(blockAlign), LenB(blockAlign)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_BITS_OFFSET), VarPtr(bits), LenB(bits)
        RtlMoveMemory ByVal (rCtx.MixFormatPtr + RIFF_WFX_CB_SIZE_OFFSET), VarPtr(cbSize), LenB(cbSize)
    End If
End Sub

'/**
' * @function VTableProc
' * @brief Reads a COM method pointer directly from an interface v-table.
' * @param pUnk COM interface pointer.
' * @param vTableIndex Zero-based v-table slot.
' * @return {Long/LongPtr} Native method pointer.
' */
#If VBA7 Then
Private Function VTableProc(ByVal pUnk As LongPtr, ByVal vTableIndex As Long) As LongPtr
    Dim pVtbl As LongPtr
    RtlMoveMemory VarPtr(pVtbl), ByVal pUnk, LenB(pVtbl)
    RtlMoveMemory VarPtr(VTableProc), ByVal (pVtbl + (vTableIndex * LenB(pVtbl))), LenB(pVtbl)
End Function
#Else
Private Function VTableProc(ByVal pUnk As Long, ByVal vTableIndex As Long) As Long
    Dim pVtbl As Long
    RtlMoveMemory VarPtr(pVtbl), ByVal pUnk, LenB(pVtbl)
    RtlMoveMemory VarPtr(VTableProc), ByVal (pVtbl + (vTableIndex * LenB(pVtbl))), LenB(pVtbl)
End Function
#End If

'/**
' * @function FastVCall0
' * @brief Invokes a COM v-table method without Variant marshaling when the hot path signature is compatible.
' */
#If VBA7 Then
Private Function FastVCall0(ByVal pUnk As LongPtr, ByVal vTableIndex As Long) As Long
#Else
Private Function FastVCall0(ByVal pUnk As Long, ByVal vTableIndex As Long) As Long
#End If
    FastVCall0 = vCall0(pUnk, vTableIndex)
End Function

'/**
' * @function FastVCall1
' * @brief Invokes a COM v-table method with one explicit argument on the optimized x64 path.
' */
#If VBA7 Then
Private Function FastVCall1(ByVal pUnk As LongPtr, ByVal vTableIndex As Long, ByVal arg0 As LongPtr) As Long
#Else
Private Function FastVCall1(ByVal pUnk As Long, ByVal vTableIndex As Long, ByVal arg0 As Long) As Long
#End If
    FastVCall1 = vCall(pUnk, vTableIndex, arg0)
End Function

'/**
' * @function FastVCall3
' * @brief Invokes a COM v-table method with three explicit arguments on the optimized x64 path.
' */
#If VBA7 Then
Private Function FastVCall3(ByVal pUnk As LongPtr, ByVal vTableIndex As Long, ByVal arg0 As LongPtr, ByVal arg1 As LongPtr, ByVal arg2 As LongPtr) As Long
#Else
Private Function FastVCall3(ByVal pUnk As Long, ByVal vTableIndex As Long, ByVal arg0 As Long, ByVal arg1 As Long, ByVal arg2 As Long) As Long
#End If
    FastVCall3 = vCall(pUnk, vTableIndex, arg0, arg1, arg2)
End Function

'/**
' * @function InitWASAPI
' * @brief Initializes the Windows Audio Session API to connect with the default sound hardware.
' * @return {Boolean} True if connected to audio driver.
' */
Private Function InitWASAPI() As Boolean
    Dim clsidEnum As GUID
    Dim iidEnum As GUID
    Dim iidAudio As GUID
    Dim iidRender As GUID
    Dim hr As Long
    Dim nChannels As Integer
    Dim bits As Integer
    
    #If Win64 Then
        Dim pNullPtr As LongLong
        Dim hnsDur As LongLong
        Dim hnsPer As LongLong
        pNullPtr = CLngLng(0)
        hnsDur = CLngLng(RIFF_DEVICE_BUFFER_MS) * CLngLng(RIFF_HNS_PER_MS)
        hnsPer = CLngLng(0)
    #Else
        #If VBA7 Then
            Dim pNullPtr As LongPtr
        #Else
            Dim pNullPtr As Long
        #End If
        Dim hnsDur As Currency
        Dim hnsPer As Currency
        pNullPtr = CLng(0)
        hnsDur = CCur(RIFF_DEVICE_BUFFER_MS)
        hnsPer = CCur(0)
    #End If
    
    IIDFromString StrPtr("{BCDE0395-E52F-467C-8E3D-C4579291692E}"), clsidEnum
    IIDFromString StrPtr(RIFF_GUID_MM_DEVICE_ENUMERATOR_CLASS), iidEnum
    IIDFromString StrPtr("{1CB9AD4C-DBFA-4c32-B178-C2F568A703B2}"), iidAudio
    IIDFromString StrPtr(RIFF_GUID_AUDIO_RENDER_CLIENT), iidRender
    
    hr = CoCreateInstance(clsidEnum, 0, CLSCTX_ALL, iidEnum, rCtx.DeviceEnumerator)
    If hr <> 0 Or rCtx.DeviceEnumerator = 0 Then
        Exit Function
    End If
    
    hr = vCall(rCtx.DeviceEnumerator, VTI_MM_DEVICE_ENUMERATOR_GET_DEFAULT_AUDIO_ENDPOINT, eRender, eConsole, VarPtr(rCtx.Device))
    If hr <> 0 Or rCtx.Device = 0 Then
        Exit Function
    End If
    
    hr = vCall(rCtx.Device, VTI_MM_DEVICE_ACTIVATE, VarPtr(iidAudio), CLSCTX_ALL, pNullPtr, VarPtr(rCtx.AudioClient))
    If hr <> 0 Or rCtx.AudioClient = 0 Then
        Exit Function
    End If
    
    hr = vCall(rCtx.AudioClient, VTI_AUDIO_CLIENT_GET_MIX_FORMAT, VarPtr(rCtx.MixFormatPtr))
    If hr <> 0 Or rCtx.MixFormatPtr = 0 Then
        Exit Function
    End If

    Dim originalMixBytes() As Byte
    Dim originalMixSize As Long
    Dim originalCbSize As Integer
    RtlMoveMemory VarPtr(originalCbSize), ByVal (rCtx.MixFormatPtr + RIFF_WFX_CB_SIZE_OFFSET), LenB(originalCbSize)
    originalMixSize = RIFF_WFX_BASE_SIZE + CLng(originalCbSize)
    If originalMixSize < RIFF_WFX_BASE_SIZE Then
        originalMixSize = RIFF_WFX_BASE_SIZE
    End If
    ReDim originalMixBytes(0 To originalMixSize - 1)
    RtlMoveMemory VarPtr(originalMixBytes(0)), ByVal rCtx.MixFormatPtr, originalMixSize

    RiffTryPromoteMixFormatToFloat32
    If Not RiffIsSharedMixFormatSupported(rCtx.MixFormatPtr) Then
        RtlMoveMemory ByVal rCtx.MixFormatPtr, VarPtr(originalMixBytes(0)), originalMixSize
    End If

    hr = vCall(rCtx.AudioClient, VTI_AUDIO_CLIENT_INITIALIZE, AUDCLNT_SHAREMODE_SHARED, 0&, hnsDur, hnsPer, rCtx.MixFormatPtr, pNullPtr)
    If hr <> 0 Then
        RtlMoveMemory ByVal rCtx.MixFormatPtr, VarPtr(originalMixBytes(0)), originalMixSize
        hr = vCall(rCtx.AudioClient, VTI_AUDIO_CLIENT_INITIALIZE, AUDCLNT_SHAREMODE_SHARED, 0&, hnsDur, hnsPer, rCtx.MixFormatPtr, pNullPtr)
    End If

    If hr <> 0 Then
        Exit Function
    End If

    RtlMoveMemory VarPtr(rCtx.sampleRate), ByVal (rCtx.MixFormatPtr + RIFF_WFX_SAMPLE_RATE_OFFSET), LenB(rCtx.sampleRate)
    RtlMoveMemory VarPtr(rCtx.AvgBytesPerSec), ByVal (rCtx.MixFormatPtr + RIFF_WFX_AVG_BYTES_OFFSET), LenB(rCtx.AvgBytesPerSec)
    RtlMoveMemory VarPtr(nChannels), ByVal (rCtx.MixFormatPtr + RIFF_WFX_CHANNELS_OFFSET), LenB(nChannels)
    RtlMoveMemory VarPtr(bits), ByVal (rCtx.MixFormatPtr + RIFF_WFX_BITS_OFFSET), LenB(bits)

    If nChannels <> RIFF_WAV_EXPORT_CHANNELS Then
        Exit Function
    End If
    If bits <> RIFF_PCM16_BITS And bits <> RIFF_FLOAT32_BITS Then
        Exit Function
    End If

    rCtx.MaxWriteFrames = (rCtx.sampleRate * 60) \ 1000
    
    hr = vCall(rCtx.AudioClient, VTI_AUDIO_CLIENT_GET_BUFFER_SIZE, VarPtr(rCtx.BufferSize))
    If hr <> 0 Then
        Exit Function
    End If
    
    hr = vCall(rCtx.AudioClient, VTI_AUDIO_CLIENT_GET_SERVICE, VarPtr(iidRender), VarPtr(rCtx.RenderClient))
    If hr <> 0 Or rCtx.RenderClient = 0 Then
        Exit Function
    End If
    
    hr = vCall0(rCtx.AudioClient, VTI_AUDIO_CLIENT_START)
    If hr <> 0 Then
        Exit Function
    End If
    
    InitWASAPI = True
End Function

'/**
' * @function ReleaseWASAPI
' * @brief Shuts down COM interfaces for the hardware connections.
' */
Private Sub ReleaseWASAPI()
    If rCtx.RenderClient <> 0 Then
        vCall0 rCtx.RenderClient, VTI_IUNKNOWN_RELEASE
        rCtx.RenderClient = 0
    End If
    
    If rCtx.AudioClient <> 0 Then
        vCall0 rCtx.AudioClient, VTI_AUDIO_CLIENT_STOP
        vCall0 rCtx.AudioClient, VTI_IUNKNOWN_RELEASE
        rCtx.AudioClient = 0
    End If
    
    If rCtx.Device <> 0 Then
        vCall0 rCtx.Device, VTI_IUNKNOWN_RELEASE
        rCtx.Device = 0
    End If
    
    If rCtx.DeviceEnumerator <> 0 Then
        vCall0 rCtx.DeviceEnumerator, VTI_IUNKNOWN_RELEASE
        rCtx.DeviceEnumerator = 0
    End If
    
    If rCtx.MixFormatPtr <> 0 Then
        CoTaskMemFree rCtx.MixFormatPtr
        rCtx.MixFormatPtr = 0
    End If
End Sub

'/**
' * @function vCall0
' * @brief Low-level C++ Virtual Table invoker for COM methods that take no arguments.
' */
#If VBA7 Then
Private Function vCall0(ByVal pUnk As LongPtr, ByVal vTableIndex As Long) As Long
#Else
Private Function vCall0(ByVal pUnk As Long, ByVal vTableIndex As Long) As Long
#End If
    Dim hrInvoke As Long
    Dim offset As Long
    Dim vRet As Variant

    offset = vTableIndex * LenB(pUnk)

    hrInvoke = DispCallFunc(pUnk, offset, CC_STDCALL, vbLong, 0, ByVal 0&, ByVal 0&, vRet)

    If hrInvoke = 0 Then
        vCall0 = CLng(vRet)
    Else
        vCall0 = hrInvoke
    End If
End Function

'/**
' * @function vCall
' * @brief Low-level C++ Virtual Table invoker to execute COM methods directly in VBA.
' * @param pUnk The pointer to the IUnknown COM interface.
' * @param vTableIndex The 0-based index of the method in the v-table.
' * @param args Array of variant parameters mapped directly to the C++ method signature.
' * @return {Long} The HRESULT returned by the COM method.
' */
#If VBA7 Then
Private Function vCall(ByVal pUnk As LongPtr, ByVal vTableIndex As Long, ParamArray args() As Variant) As Long
#Else
Private Function vCall(ByVal pUnk As Long, ByVal vTableIndex As Long, ParamArray args() As Variant) As Long
#End If
    Dim hrInvoke As Long
    Dim i As Long
    Dim argCount As Long
    Dim offset As Long
    Dim vRet As Variant
    
    offset = vTableIndex * LenB(pUnk)
    
    argCount = UBound(args) - LBound(args) + 1
    If argCount < 0 Then
        argCount = 0
    End If
    
    If argCount > 0 Then
        #If VBA7 Then
            Dim vArgs() As Variant
            Dim vTypes() As Integer
            Dim pArgs() As LongPtr
        #Else
            Dim vArgs() As Variant
            Dim vTypes() As Integer
            Dim pArgs() As Long
        #End If
        
        ReDim vArgs(0 To argCount - 1)
        ReDim vTypes(0 To argCount - 1)
        ReDim pArgs(0 To argCount - 1)
        
        For i = 0 To argCount - 1
            vArgs(i) = args(i)
            vTypes(i) = VarType(vArgs(i))
            pArgs(i) = VarPtr(vArgs(i))
        Next i
        
        hrInvoke = DispCallFunc(pUnk, offset, CC_STDCALL, vbLong, argCount, vTypes(0), pArgs(0), vRet)
    Else
        hrInvoke = DispCallFunc(pUnk, offset, CC_STDCALL, vbLong, 0, ByVal 0&, ByVal 0&, vRet)
    End If
    
    If hrInvoke = 0 Then
        vCall = CLng(vRet)
    Else
        vCall = hrInvoke
    End If
End Function
