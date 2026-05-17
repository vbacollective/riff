Attribute VB_Name = "Riff"
'/**
' * Riff - Audio Engine (Studio DSP Edition)
' * @description A high-performance, COM-based WASAPI audio engine for VBA (x86/x64 compatible).
' * Contains advanced Array Chunking for zero-latency mixing, Polyphony, and a full
' * Studio DSP Pipeline featuring Reverb, Chorus, Flanger, Compressor, 3-Band EQ, Bitcrusher,
' * RingMod, AutoPan, Delay, Oscillators, In-Memory Loading, Buses, and Peak Meters.
' * @author UesleiDev
' * @version 1.0.0
' */

Option Explicit
Option Private Module

' ========================================================================================
' API DECLARATIONS: MEMORY, WINDOWS, AND COM INTEROP
' ========================================================================================

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
    
    '/** @description Creates an IStream from a Byte Array in memory. */
    Private Declare PtrSafe Function SHCreateMemStream Lib "shlwapi.dll" (ByRef pInit As Any, ByVal cbInit As Long) As LongPtr
    
    '/** @description Wraps an IStream into an IMFByteStream. */
    Private Declare PtrSafe Function MFCreateMFByteStreamOnStream Lib "mfplat.dll" (ByVal pStream As LongPtr, ByRef ppByteStream As LongPtr) As Long
    
    '/** @description Creates an MF Source Reader directly from an IMFByteStream. */
    Private Declare PtrSafe Function MFCreateSourceReaderFromByteStream Lib "mfreadwrite.dll" (ByVal pByteStream As LongPtr, ByVal pAttributes As LongPtr, ByRef ppSourceReader As LongPtr) As Long
    
    '/** @description Increases system timer resolution. */
    Private Declare PtrSafe Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
    
    '/** @description Restores system timer resolution. */
    Private Declare PtrSafe Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
    
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
        SampleRate As Long
        AvgBytesPerSec As Long
        DeviceEnumerator As LongPtr
        Device As LongPtr
        AudioClient As LongPtr
        RenderClient As LongPtr
        MixFormatPtr As LongPtr
        BufferSize As Long
        ThunkTimerCB As LongPtr
        TimerID As LongPtr
        Buffers(0 To 63) As RiffBuffer
        Buses(0 To 7) As Single
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
    
    '/** @description Creates an IStream from a Byte Array in memory. */
    Private Declare Function SHCreateMemStream Lib "shlwapi.dll" (ByRef pInit As Any, ByVal cbInit As Long) As Long
    
    '/** @description Wraps an IStream into an IMFByteStream. */
    Private Declare Function MFCreateMFByteStreamOnStream Lib "mfplat.dll" (ByVal pStream As Long, ByRef ppByteStream As Long) As Long
    
    '/** @description Creates an MF Source Reader directly from an IMFByteStream. */
    Private Declare Function MFCreateSourceReaderFromByteStream Lib "mfreadwrite.dll" (ByVal pByteStream As Long, ByVal pAttributes As Long, ByRef ppSourceReader As Long) As Long
    
    '/** @description Increases system timer resolution. */
    Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
    
    '/** @description Restores system timer resolution. */
    Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
    
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
        SampleRate As Long
        AvgBytesPerSec As Long
        DeviceEnumerator As Long
        Device As Long
        AudioClient As Long
        RenderClient As Long
        MixFormatPtr As Long
        BufferSize As Long
        ThunkTimerCB As Long
        TimerID As Long
        Buffers(0 To 63) As RiffBuffer
        Buses(0 To 7) As Single
    End Type
#End If

'/**
' * @struct RiffVoice
' * @brief Polyphonic playback channel containing a full Studio DSP Pipeline matrix.
' */
Private Type RiffVoice
    Active As Boolean
    Playing As Boolean
    Paused As Boolean
    busID As Long
    BufferIndex As Long
    
    ' --- Generators ---
    IsOscillator As Boolean
    OscType As Long
    OscFreq As Single
    OscPhase As Double
    
    ' --- Master Metrics ---
    PeakL As Single
    PeakR As Single
    
    ' --- Basic Control ---
    Position As Double
    Volume As Single
    Pitch As Double
    Pan As Single
    
    ' --- DSP Pipeline States ---
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
    
    LowPass As Single
    HighPass As Single
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
    
    DelayTime As Single
    DelayFeedback As Single
    DelayMix As Single
    
    CompThreshold As Single
    CompRatio As Single
    CompEnv As Single
    
    RingWritePos As Long
    
    ' --- Loop and Fade Control ---
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

'/** @description Mathematical constant Pi. */
Private Const PI As Double = 3.14159265358979

'/** @description Mathematical constant 2 * Pi. */
Private Const PI2 As Double = 6.28318530717958

'/** @description Global state holding hardware info and context. */
Private rCtx As RiffContext

'/** @description Pool of 32 polyphonic voices for audio playback. */
Private rVoices(0 To 31) As RiffVoice

'/**
' * @description Contiguous global ring buffer array for spatial effects.
' * Solves the Column-Major 2D wipe issue by providing 1D sequential access.
' */
Private rRingBuf() As Single

' ========================================================================================
' PUBLIC API: INITIALIZATION AND TEARDOWN
' ========================================================================================

'/**
' * @function RiffOpen
' * @brief Initializes the WASAPI Audio Engine, Media Foundation, and background Timers.
' * @return {Boolean} True if initialization was successful.
' */
Public Function RiffOpen() As Boolean
    If rCtx.Initialized Then
        RiffOpen = True
        Exit Function
    End If
    
    rCtx.MagicCookie = &H52494646
    rCtx.MasterVolume = 1!
    rCtx.MasterPeakL = 0!
    rCtx.MasterPeakR = 0!
    
    Dim i As Long
    For i = 0 To 7
        rCtx.Buses(i) = 1!
    Next i
    
    If Not InitThunks() Then
        Exit Function
    End If
    
    If MFStartup(MF_VERSION, 0) <> 0 Then
        Exit Function
    End If
    
    If Not InitWASAPI() Then
        ReleaseWASAPI
        MFShutdown
        FreeThunks
        Exit Function
    End If
    
    ' Allocate 32 contiguous buffers (192,000 samples each) in a strictly 1D array.
    ' This prevents the VBA Column-Major alignment issue on RtlZeroMemory.
    ReDim rRingBuf(0 To (32 * 192000) - 1)
    
    timeBeginPeriod 1
    
    rCtx.TimerID = SetTimer(0, 0, 15, rCtx.ThunkTimerCB)
    
    If rCtx.TimerID = 0 Then
        timeEndPeriod 1
        ReleaseWASAPI
        MFShutdown
        FreeThunks
        Exit Function
    End If
    
    rCtx.Initialized = True
    RiffOpen = True
End Function

'/**
' * @function RiffClose
' * @brief Safely shuts down the audio engine, frees memory, and releases COM objects.
' */
Public Sub RiffClose()
    If Not rCtx.Initialized Then
        Exit Sub
    End If
    
    rCtx.MagicCookie = 0
    
    If rCtx.TimerID <> 0 Then
        KillTimer 0, rCtx.TimerID
        rCtx.TimerID = 0
        timeEndPeriod 1
    End If
    
    Dim i As Long
    For i = 0 To 31
        rVoices(i).Active = False
    Next i
    
    For i = 0 To 63
        If rCtx.Buffers(i).Active Then
            RiffUnload i
        End If
    Next i
    
    Erase rRingBuf
    
    ReleaseWASAPI
    MFShutdown
    FreeThunks
    
    rCtx.Initialized = False
End Sub

' ========================================================================================
' PUBLIC API: GLOBAL SETTINGS & ASSET MANAGEMENT
' ========================================================================================

'/**
' * @property RiffIsInitialized
' * @brief Returns True if the audio engine is running.
' */
Public Property Get RiffIsInitialized() As Boolean
    RiffIsInitialized = rCtx.Initialized
End Property

'/**
' * @property RiffMasterVolume
' * @brief Global master volume for the engine (0.0 to 1.0).
' */
Public Property Get RiffMasterVolume() As Single
    If Not rCtx.Initialized Then
        Exit Property
    End If
    
    RiffMasterVolume = rCtx.MasterVolume
End Property
Public Property Let RiffMasterVolume(ByVal value As Single)
    If Not rCtx.Initialized Then
        Exit Property
    End If
    
    If value < 0! Then
        value = 0!
    End If
    If value > 1! Then
        value = 1!
    End If
    
    rCtx.MasterVolume = value
End Property

'/**
' * @property RiffBusVolume
' * @brief Sets the volume multiplier for a specific Audio Bus (0 to 7).
' * @param busID The target bus index.
' */
Public Property Get RiffBusVolume(ByVal busID As Long) As Single
    If Not rCtx.Initialized Then
        Exit Property
    End If
    
    If busID < 0 Then
        busID = 0
    End If
    If busID > 7 Then
        busID = 7
    End If
    
    RiffBusVolume = rCtx.Buses(busID)
End Property
Public Property Let RiffBusVolume(ByVal busID As Long, ByVal value As Single)
    If Not rCtx.Initialized Then
        Exit Property
    End If
    
    If busID < 0 Then
        busID = 0
    End If
    If busID > 7 Then
        busID = 7
    End If
    
    If value < 0! Then
        value = 0!
    End If
    If value > 2! Then
        value = 2!
    End If
    
    rCtx.Buses(busID) = value
End Property

'/**
' * @function RiffMasterGetPeak
' * @brief Retrieves the absolute peak amplitude of the master output for VU Meters.
' * @param peakLeft Variable to store the left channel peak.
' * @param peakRight Variable to store the right channel peak.
' */
Public Sub RiffMasterGetPeak(ByRef peakLeft As Single, ByRef peakRight As Single)
    If Not rCtx.Initialized Then
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
    
    If Not rCtx.Initialized Then
        Exit Function
    End If
    
    Dim slot As Long
    Dim i As Long
    slot = -1
    
    For i = 0 To 63
        If Not rCtx.Buffers(i).Active Then
            slot = i
            Exit For
        End If
    Next i
    
    If slot = -1 Then
        Exit Function
    End If
    
    If Dir(filePath) = "" Then
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
    IIDFromString StrPtr("{7632CB14-D379-4770-AE7D-EA24154D9298}"), guidResample
    
    hr = MFCreateAttributes(pAttributes, 1)
    
    If hr = 0 And pAttributes <> 0 Then
        vCall pAttributes, 21, VarPtr(guidResample), 1&
        hr = MFCreateSourceReaderFromURL(StrPtr(filePath), pAttributes, pReader)
        vCall pAttributes, 2
    Else
        hr = MFCreateSourceReaderFromURL(StrPtr(filePath), 0, pReader)
    End If
    
    If hr <> 0 Or pReader = 0 Then
        Exit Function
    End If
    
    RiffLoad = CoreProcessSourceReader(pReader, slot)
End Function

'/**
' * @function RiffLoadFromMemory
' * @brief Decodes audio data directly from a Byte Array without requiring disk I/O.
' * @param audioData Byte array containing the binary audio file.
' * @return {Long} Buffer handle (0-63), or -1 if failed.
' */
Public Function RiffLoadFromMemory(ByRef audioData() As Byte) As Long
    RiffLoadFromMemory = -1
    
    If Not rCtx.Initialized Then
        Exit Function
    End If
    
    Dim slot As Long
    Dim i As Long
    slot = -1
    
    For i = 0 To 63
        If Not rCtx.Buffers(i).Active Then
            slot = i
            Exit For
        End If
    Next i
    
    If slot = -1 Then
        Exit Function
    End If
    
    #If VBA7 Then
        Dim pStream As LongPtr
        Dim pByteStream As LongPtr
        Dim pReader As LongPtr
        Dim pAttributes As LongPtr
    #Else
        Dim pStream As Long
        Dim pByteStream As Long
        Dim pReader As Long
        Dim pAttributes As Long
    #End If
    
    Dim cbSize As Long
    cbSize = UBound(audioData) - LBound(audioData) + 1
    
    If cbSize <= 0 Then
        Exit Function
    End If
    
    pStream = SHCreateMemStream(audioData(LBound(audioData)), cbSize)
    If pStream = 0 Then
        Exit Function
    End If
    
    Dim hr As Long
    hr = MFCreateMFByteStreamOnStream(pStream, pByteStream)
    
    If hr <> 0 Or pByteStream = 0 Then
        vCall pStream, 2
        Exit Function
    End If
    
    Dim guidResample As GUID
    IIDFromString StrPtr("{7632CB14-D379-4770-AE7D-EA24154D9298}"), guidResample
    
    hr = MFCreateAttributes(pAttributes, 1)
    
    If hr = 0 And pAttributes <> 0 Then
        vCall pAttributes, 21, VarPtr(guidResample), 1&
        hr = MFCreateSourceReaderFromByteStream(pByteStream, pAttributes, pReader)
        vCall pAttributes, 2
    Else
        hr = MFCreateSourceReaderFromByteStream(pByteStream, 0, pReader)
    End If
    
    If hr = 0 And pReader <> 0 Then
        RiffLoadFromMemory = CoreProcessSourceReader(pReader, slot)
    End If
    
    If pByteStream <> 0 Then
        vCall pByteStream, 2
    End If
    
    If pStream <> 0 Then
        vCall pStream, 2
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
        Dim llTime As LongLong
        pNullPtr = CLngPtr(0)
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
    
    vCall pReader, 4, MF_SOURCE_READER_ALL_STREAMS, 0&
    vCall pReader, 4, MF_SOURCE_READER_FIRST_AUDIO_STREAM, 1&
    
    hr = MFCreateMediaType(pPartialType)
    If hr = 0 And pPartialType <> 0 Then
        Dim wfx_cbSize As Integer
        RtlMoveMemory VarPtr(wfx_cbSize), ByVal (rCtx.MixFormatPtr + 16), 2
        
        hr = MFInitMediaTypeFromWaveFormatEx(pPartialType, rCtx.MixFormatPtr, 18 + CLng(wfx_cbSize))
        If hr = 0 Then
            vCall pReader, 7, MF_SOURCE_READER_FIRST_AUDIO_STREAM, pNullPtr, pPartialType
        End If
        vCall pPartialType, 2
    End If
    
    Dim pSz As Long
    #If Win64 Then
        pSz = 8
    #Else
        pSz = 4
    #End If

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
    
    currentCap = 1048576 * 50
    tempPtr = VirtualAlloc(0, currentCap, MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)
    
    If tempPtr = 0 Then
        vCall pReader, 2
        Exit Function
    End If
    
    totalSize = 0
    Do
        pSample = 0
        hrInvoke = DispCallFunc(pReader, 9 * pSz, CC_STDCALL, vbLong, 6, rTypes(0), rPtrs(0), vRet)
        
        If hrInvoke <> 0 Then
            hr = hrInvoke
        Else
            hr = CLng(vRet)
        End If
        
        If hr <> 0 Or (dwFlags And 2) <> 0 Then
            Exit Do
        End If
        
        If pSample <> 0 Then
            hrInvoke = DispCallFunc(pSample, 41 * pSz, CC_STDCALL, vbLong, 1, cTypes(0), cPtrs(0), vRet)
            
            If hrInvoke <> 0 Then
                hr = hrInvoke
            Else
                hr = CLng(vRet)
            End If
            
            If hr = 0 And pBuffer <> 0 Then
                hrInvoke = DispCallFunc(pBuffer, 3 * pSz, CC_STDCALL, vbLong, 3, lTypes(0), lPtrs(0), vRet)
                
                If hrInvoke <> 0 Then
                    hr = hrInvoke
                Else
                    hr = CLng(vRet)
                End If
                
                If hr = 0 And pAudioData <> 0 And cbLen > 0 Then
                    If totalSize + cbLen > currentCap Then
                        currentCap = (totalSize + cbLen) * 2
                        newPtr = VirtualAlloc(0, currentCap, MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)
                        
                        If newPtr <> 0 Then
                            RtlMoveMemory ByVal newPtr, ByVal tempPtr, totalSize
                            VirtualFree tempPtr, 0, MEM_RELEASE
                            tempPtr = newPtr
                        End If
                    End If
                    
                    If tempPtr <> 0 Then
                        RtlMoveMemory ByVal (tempPtr + totalSize), ByVal pAudioData, cbLen
                        totalSize = totalSize + cbLen
                    End If
                    
                    DispCallFunc pBuffer, 4 * pSz, CC_STDCALL, vbLong, 0, ByVal 0&, ByVal 0&, vRet
                End If
                
                DispCallFunc pBuffer, 2 * pSz, CC_STDCALL, vbLong, 0, ByVal 0&, ByVal 0&, vRet
            End If
            
            DispCallFunc pSample, 2 * pSz, CC_STDCALL, vbLong, 0, ByVal 0&, ByVal 0&, vRet
        End If
    Loop
    
    vCall pReader, 2
    
    If totalSize > 0 Then
        rCtx.Buffers(slot).BufferPtr = VirtualAlloc(0, totalSize, MEM_COMMIT Or MEM_RESERVE, PAGE_READWRITE)
        
        If rCtx.Buffers(slot).BufferPtr <> 0 Then
            RtlMoveMemory ByVal rCtx.Buffers(slot).BufferPtr, ByVal tempPtr, totalSize
            rCtx.Buffers(slot).BufferLen = totalSize
            rCtx.Buffers(slot).Active = True
            CoreProcessSourceReader = slot
        End If
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
    If Not rCtx.Initialized Then
        Exit Sub
    End If
    If bufferHandle < 0 Or bufferHandle > 63 Then
        Exit Sub
    End If
    If Not rCtx.Buffers(bufferHandle).Active Then
        Exit Sub
    End If
    
    Dim i As Long
    For i = 0 To 31
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
    If Not rCtx.Initialized Or bufferHandle < 0 Or bufferHandle > 63 Then
        Exit Property
    End If
    If Not rCtx.Buffers(bufferHandle).Active Or rCtx.AvgBytesPerSec = 0 Then
        Exit Property
    End If
    
    RiffBufferDurationSec = CSng(rCtx.Buffers(bufferHandle).BufferLen) / CSng(rCtx.AvgBytesPerSec)
End Property

' ========================================================================================
' PUBLIC API: PLAYBACK & VOICE ACTIONS
' ========================================================================================

'/**
' * @function RiffPlay
' * @brief Plays a pre-loaded static audio buffer on an available polyphonic voice channel.
' * @param bufferHandle The target buffer.
' * @return {Long} Voice handle (0-31), or -1 if all voices are occupied.
' */
Public Function RiffPlay(ByVal bufferHandle As Long) As Long
    RiffPlay = -1
    
    If Not rCtx.Initialized Then
        Exit Function
    End If
    If bufferHandle < 0 Or bufferHandle > 63 Then
        Exit Function
    End If
    If Not rCtx.Buffers(bufferHandle).Active Then
        Exit Function
    End If
    
    Dim voiceSlot As Long
    voiceSlot = InternalGetFreeVoice()
    
    If voiceSlot = -1 Then
        Exit Function
    End If
    
    InternalResetVoiceDSP voiceSlot
    
    rVoices(voiceSlot).IsOscillator = False
    rVoices(voiceSlot).BufferIndex = bufferHandle
    rVoices(voiceSlot).Position = 0#
    rVoices(voiceSlot).loopEnd = CDbl(rCtx.Buffers(bufferHandle).BufferLen)
    rVoices(voiceSlot).Playing = True
    rVoices(voiceSlot).Active = True
    
    RiffPlay = voiceSlot
End Function

'/**
' * @function RiffPlayOscillator
' * @brief Generates and plays a synthesized waveform instantly via math.
' * @param waveType 0=Sine, 1=Square, 2=Sawtooth, 3=Noise.
' * @param frequencyHz Pitch frequency in Hz.
' * @return {Long} Voice handle (0-31), or -1 if all voices are occupied.
' */
Public Function RiffPlayOscillator(ByVal waveType As Long, ByVal frequencyHz As Single) As Long
    RiffPlayOscillator = -1
    
    If Not rCtx.Initialized Then
        Exit Function
    End If
    
    Dim voiceSlot As Long
    voiceSlot = InternalGetFreeVoice()
    
    If voiceSlot = -1 Then
        Exit Function
    End If
    
    InternalResetVoiceDSP voiceSlot
    
    If waveType < 0 Then
        waveType = 0
    End If
    If waveType > 3 Then
        waveType = 3
    End If
    
    If frequencyHz < 1! Then
        frequencyHz = 440!
    End If
    
    rVoices(voiceSlot).IsOscillator = True
    rVoices(voiceSlot).OscType = waveType
    rVoices(voiceSlot).OscFreq = frequencyHz
    rVoices(voiceSlot).OscPhase = 0#
    rVoices(voiceSlot).BufferIndex = -1
    
    rVoices(voiceSlot).Playing = True
    rVoices(voiceSlot).Active = True
    
    RiffPlayOscillator = voiceSlot
End Function

'/**
' * @function InternalGetFreeVoice
' * @brief Scans for an inactive voice channel.
' * @return {Long} Voice index, or -1.
' */
Private Function InternalGetFreeVoice() As Long
    Dim i As Long
    InternalGetFreeVoice = -1
    
    For i = 0 To 31
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
    rVoices(slot).Volume = 1!
    rVoices(slot).Pitch = 1#
    rVoices(slot).Pan = 0!
    rVoices(slot).Paused = False
    rVoices(slot).busID = 0
    rVoices(slot).PeakL = 0!
    rVoices(slot).PeakR = 0!
    
    rVoices(slot).Distortion = 1!
    rVoices(slot).LowPass = 1!
    rVoices(slot).HighPass = 0!
    rVoices(slot).FilterStateL = 0!
    rVoices(slot).FilterStateR = 0!
    rVoices(slot).FilterStateHP_L = 0!
    rVoices(slot).FilterStateHP_R = 0!
    rVoices(slot).StereoWidth = 1!
    
    rVoices(slot).EqBass = 1!
    rVoices(slot).EqMid = 1!
    rVoices(slot).EqTreble = 1!
    rVoices(slot).EqStateLowL = 0!
    rVoices(slot).EqStateLowR = 0!
    rVoices(slot).EqStateHighL = 0!
    rVoices(slot).EqStateHighR = 0!
    
    rVoices(slot).CompThreshold = 1!
    rVoices(slot).CompRatio = 1!
    rVoices(slot).CompEnv = 0.0001! ' Prevent division by zero mathematically
    
    rVoices(slot).BitcrushSteps = 0!
    rVoices(slot).BitcrushDownsample = 1
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
    rVoices(slot).ChorusRate = 1.5!
    rVoices(slot).ChorusPhase = 0#
    
    rVoices(slot).FlangerRate = 0.5!
    rVoices(slot).FlangerDepth = 0!
    rVoices(slot).FlangerFeedback = 0!
    rVoices(slot).FlangerPhase = 0#
    
    rVoices(slot).ReverbMix = 0!
    rVoices(slot).ReverbTime = 0.5!
    rVoices(slot).RevTap1 = (Int(0.029 * rCtx.SampleRate) \ 2) * 2
    rVoices(slot).RevTap2 = (Int(0.043 * rCtx.SampleRate) \ 2) * 2
    rVoices(slot).RevTap3 = (Int(0.073 * rCtx.SampleRate) \ 2) * 2
    rVoices(slot).RevTap4 = (Int(0.097 * rCtx.SampleRate) \ 2) * 2
    
    rVoices(slot).DelayTime = 0!
    rVoices(slot).DelayFeedback = 0!
    rVoices(slot).DelayMix = 0!
    rVoices(slot).RingWritePos = 0
    
    rVoices(slot).Looping = False
    rVoices(slot).loopStart = 0#
    rVoices(slot).fadeState = 0
    
    ' Zero the exact 1D mapped block associated with this voice
    RtlZeroMemory VarPtr(rRingBuf(slot * 192000)), 192000 * 4
End Sub

'/**
' * @function RiffPause
' * @brief Pauses a specific voice without killing it.
' * @param voiceHandle The active voice.
' */
Public Sub RiffPause(ByVal voiceHandle As Long)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
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
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Sub
    End If
    
    If rVoices(voiceHandle).Active Then
        rVoices(voiceHandle).Paused = False
    End If
End Sub

'/**
' * @function RiffStop
' * @brief Immediately halts and frees a playing voice.
' * @param voiceHandle The active voice.
' */
Public Sub RiffStop(ByVal voiceHandle As Long)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
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
    Dim i As Long
    For i = 0 To 31
        rVoices(i).Playing = False
        rVoices(i).Active = False
    Next i
End Sub

'/**
' * @function RiffFadeIn
' * @brief Smoothly fades in a voice over the specified duration.
' * @param voiceHandle The active voice.
' * @param durationSec The time in seconds to reach full volume.
' */
Public Sub RiffFadeIn(ByVal voiceHandle As Long, ByVal durationSec As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Sub
    End If
    If durationSec <= 0! Then
        Exit Sub
    End If
    
    rVoices(voiceHandle).FadeFramesTotal = CLng(durationSec * rCtx.SampleRate)
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
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Sub
    End If
    If durationSec <= 0! Then
        Exit Sub
    End If
    
    rVoices(voiceHandle).FadeFramesTotal = CLng(durationSec * rCtx.SampleRate)
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
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Sub
    End If
    If rVoices(voiceHandle).IsOscillator Then
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
    
    If eByte > rCtx.Buffers(bufHandle).BufferLen Then
        eByte = rCtx.Buffers(bufHandle).BufferLen
    End If
    
    Dim nBlockAlign As Integer
    RtlMoveMemoryToInteger nBlockAlign, ByVal (rCtx.MixFormatPtr + 12), 2
    
    Dim align As Long
    align = CLng(nBlockAlign)
    
    sByte = CDbl((CLng(sByte) \ align) * align)
    eByte = CDbl((CLng(eByte) \ align) * align)
    
    rVoices(voiceHandle).loopStart = sByte
    rVoices(voiceHandle).loopEnd = eByte
End Sub

' ========================================================================================
' PUBLIC API: VOICE PROPERTIES (DSP EFFECTS & STATE)
' ========================================================================================

'/**
' * @property RiffVoiceIsPlaying
' * @brief Checks if a voice is actively processing audio.
' */
Public Property Get RiffVoiceIsPlaying(ByVal voiceHandle As Long) As Boolean
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceIsPlaying = (rVoices(voiceHandle).Active And rVoices(voiceHandle).Playing)
End Property

'/**
' * @property RiffVoiceIsPaused
' * @brief Checks if a voice is currently paused.
' */
Public Property Get RiffVoiceIsPaused(ByVal voiceHandle As Long) As Boolean
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceIsPaused = rVoices(voiceHandle).Paused
End Property

'/**
' * @property RiffVoiceBus
' * @brief Determines which Audio Bus this voice routes to (0 to 7).
' */
Public Property Get RiffVoiceBus(ByVal voiceHandle As Long) As Long
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceBus = rVoices(voiceHandle).busID
End Property
Public Property Let RiffVoiceBus(ByVal voiceHandle As Long, ByVal value As Long)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    
    If value < 0 Then
        value = 0
    End If
    If value > 7 Then
        value = 7
    End If
    
    rVoices(voiceHandle).busID = value
End Property

'/**
' * @function RiffVoiceGetPeak
' * @brief Retrieves the instantaneous Peak Amplitude for this voice's VU Meters.
' * @param peakLeft Variable to store the left channel peak.
' * @param peakRight Variable to store the right channel peak.
' */
Public Sub RiffVoiceGetPeak(ByVal voiceHandle As Long, ByRef peakLeft As Single, ByRef peakRight As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
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
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceLoop = rVoices(voiceHandle).Looping
End Property
Public Property Let RiffVoiceLoop(ByVal voiceHandle As Long, ByVal value As Boolean)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    rVoices(voiceHandle).Looping = value
End Property

'/**
' * @property RiffVoicePositionSec
' * @brief Current playback position in seconds.
' */
Public Property Get RiffVoicePositionSec(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Or rCtx.AvgBytesPerSec = 0 Then
        Exit Property
    End If
    RiffVoicePositionSec = CSng(rVoices(voiceHandle).Position) / CSng(rCtx.AvgBytesPerSec)
End Property
Public Property Let RiffVoicePositionSec(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
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
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceVolume = rVoices(voiceHandle).Volume
End Property
Public Property Let RiffVoiceVolume(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    
    If value < 0! Then
        value = 0!
    End If
    If value > 1! Then
        value = 1!
    End If
    
    rVoices(voiceHandle).Volume = value
End Property

'/**
' * @property RiffVoicePitch
' * @brief Speed/Pitch modifier (1.0 = Normal, 2.0 = Double speed/octave higher).
' */
Public Property Get RiffVoicePitch(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoicePitch = CSng(rVoices(voiceHandle).Pitch)
End Property
Public Property Let RiffVoicePitch(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    
    If value <= 0.1! Then
        value = 0.1!
    End If
    
    rVoices(voiceHandle).Pitch = CDbl(value)
End Property

'/**
' * @property RiffVoicePan
' * @brief Stereo panning (-1.0 = Left, 0.0 = Center, 1.0 = Right).
' */
Public Property Get RiffVoicePan(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoicePan = rVoices(voiceHandle).Pan
End Property
Public Property Let RiffVoicePan(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    
    If value < -1! Then
        value = -1!
    End If
    If value > 1! Then
        value = 1!
    End If
    
    rVoices(voiceHandle).Pan = value
End Property

' ========================================================================================
' DSP FILTERS API
' ========================================================================================

'/**
' * @property RiffVoiceBitDepth
' * @brief Simulates retro console audio by quantizing amplitude bits.
' */
Public Property Get RiffVoiceBitDepth(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    
    If rVoices(voiceHandle).BitcrushSteps = 0! Then
        RiffVoiceBitDepth = 32!
    Else
        RiffVoiceBitDepth = Log(rVoices(voiceHandle).BitcrushSteps) / Log(2)
    End If
End Property
Public Property Let RiffVoiceBitDepth(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    
    If value >= 32! Then
        rVoices(voiceHandle).BitcrushSteps = 0!
        Exit Property
    End If
    If value < 2! Then
        value = 2!
    End If
    
    rVoices(voiceHandle).BitcrushSteps = 2 ^ value
End Property

'/**
' * @property RiffVoiceSampleRateReduction
' * @brief Creates robotic artifacts by holding samples for N frames.
' */
Public Property Get RiffVoiceSampleRateReduction(ByVal voiceHandle As Long) As Long
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceSampleRateReduction = rVoices(voiceHandle).BitcrushDownsample
End Property
Public Property Let RiffVoiceSampleRateReduction(ByVal voiceHandle As Long, ByVal value As Long)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 1 Then
        value = 1
    End If
    rVoices(voiceHandle).BitcrushDownsample = value
End Property

'/**
' * @property RiffVoiceRingModFreq
' * @brief The frequency in Hz for the Ring Modulator oscillator.
' */
Public Property Get RiffVoiceRingModFreq(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceRingModFreq = rVoices(voiceHandle).RingModFreq
End Property
Public Property Let RiffVoiceRingModFreq(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
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
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceRingModMix = rVoices(voiceHandle).RingModMix
End Property
Public Property Let RiffVoiceRingModMix(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 1! Then
        value = 1!
    End If
    rVoices(voiceHandle).RingModMix = value
End Property

'/**
' * @property RiffVoiceAutoPanRate
' * @brief Speed of the automatic panning LFO in Hz.
' */
Public Property Get RiffVoiceAutoPanRate(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceAutoPanRate = rVoices(voiceHandle).AutoPanRate
End Property
Public Property Let RiffVoiceAutoPanRate(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 20! Then
        value = 20!
    End If
    rVoices(voiceHandle).AutoPanRate = value
End Property

'/**
' * @property RiffVoiceAutoPanDepth
' * @brief Intensity of the auto-panning effect.
' */
Public Property Get RiffVoiceAutoPanDepth(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceAutoPanDepth = rVoices(voiceHandle).AutoPanDepth
End Property
Public Property Let RiffVoiceAutoPanDepth(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 1! Then
        value = 1!
    End If
    rVoices(voiceHandle).AutoPanDepth = value
End Property

'/**
' * @property RiffVoiceEqBass
' * @brief Low frequency shelf gain (1.0 is flat).
' */
Public Property Get RiffVoiceEqBass(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceEqBass = rVoices(voiceHandle).EqBass
End Property
Public Property Let RiffVoiceEqBass(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 5! Then
        value = 5!
    End If
    rVoices(voiceHandle).EqBass = value
End Property

'/**
' * @property RiffVoiceEqMid
' * @brief Mid frequency band gain (1.0 is flat).
' */
Public Property Get RiffVoiceEqMid(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceEqMid = rVoices(voiceHandle).EqMid
End Property
Public Property Let RiffVoiceEqMid(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 5! Then
        value = 5!
    End If
    rVoices(voiceHandle).EqMid = value
End Property

'/**
' * @property RiffVoiceEqTreble
' * @brief High frequency shelf gain (1.0 is flat).
' */
Public Property Get RiffVoiceEqTreble(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceEqTreble = rVoices(voiceHandle).EqTreble
End Property
Public Property Let RiffVoiceEqTreble(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 5! Then
        value = 5!
    End If
    rVoices(voiceHandle).EqTreble = value
End Property

'/**
' * @property RiffVoiceCompressorThreshold
' * @brief Volume level at which the compressor starts reducing gain.
' */
Public Property Get RiffVoiceCompressorThreshold(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceCompressorThreshold = rVoices(voiceHandle).CompThreshold
End Property
Public Property Let RiffVoiceCompressorThreshold(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value <= 0! Then
        value = 0.01!
    End If
    If value > 1! Then
        value = 1!
    End If
    rVoices(voiceHandle).CompThreshold = value
End Property

'/**
' * @property RiffVoiceCompressorRatio
' * @brief Amount of gain reduction applied when signal exceeds the threshold.
' */
Public Property Get RiffVoiceCompressorRatio(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceCompressorRatio = rVoices(voiceHandle).CompRatio
End Property
Public Property Let RiffVoiceCompressorRatio(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 1! Then
        value = 1!
    End If
    If value > 20! Then
        value = 20!
    End If
    rVoices(voiceHandle).CompRatio = value
End Property

'/**
' * @property RiffVoiceFlangerDepth
' * @brief Blend amount for the Flanger effect.
' */
Public Property Get RiffVoiceFlangerDepth(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceFlangerDepth = rVoices(voiceHandle).FlangerDepth
End Property
Public Property Let RiffVoiceFlangerDepth(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 1! Then
        value = 1!
    End If
    rVoices(voiceHandle).FlangerDepth = value
End Property

'/**
' * @property RiffVoiceFlangerRate
' * @brief Sweep rate of the Flanger in Hz.
' */
Public Property Get RiffVoiceFlangerRate(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceFlangerRate = rVoices(voiceHandle).FlangerRate
End Property
Public Property Let RiffVoiceFlangerRate(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0.1! Then
        value = 0.1!
    End If
    If value > 10! Then
        value = 10!
    End If
    rVoices(voiceHandle).FlangerRate = value
End Property

'/**
' * @property RiffVoiceFlangerFeedback
' * @brief Resonance intensity of the Flanger effect.
' */
Public Property Get RiffVoiceFlangerFeedback(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceFlangerFeedback = rVoices(voiceHandle).FlangerFeedback
End Property
Public Property Let RiffVoiceFlangerFeedback(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 0.95! Then
        value = 0.95!
    End If
    rVoices(voiceHandle).FlangerFeedback = value
End Property

'/**
' * @property RiffVoiceDistortion
' * @brief Digital clipping multiplier.
' */
Public Property Get RiffVoiceDistortion(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceDistortion = rVoices(voiceHandle).Distortion
End Property
Public Property Let RiffVoiceDistortion(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 1! Then
        value = 1!
    End If
    rVoices(voiceHandle).Distortion = value
End Property

'/**
' * @property RiffVoiceLowPass
' * @brief Muffles the audio by filtering high frequencies.
' */
Public Property Get RiffVoiceLowPass(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceLowPass = rVoices(voiceHandle).LowPass
End Property
Public Property Let RiffVoiceLowPass(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value <= 0! Then
        value = 0.01!
    End If
    If value > 1! Then
        value = 1!
    End If
    rVoices(voiceHandle).LowPass = value
End Property

'/**
' * @property RiffVoiceHighPass
' * @brief Thins out the audio by filtering low frequencies.
' */
Public Property Get RiffVoiceHighPass(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceHighPass = rVoices(voiceHandle).HighPass
End Property
Public Property Let RiffVoiceHighPass(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 0.99! Then
        value = 0.99!
    End If
    rVoices(voiceHandle).HighPass = value
End Property

'/**
' * @property RiffVoiceStereoWidth
' * @brief Adjusts the perceived width of the stereo field.
' */
Public Property Get RiffVoiceStereoWidth(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceStereoWidth = rVoices(voiceHandle).StereoWidth
End Property
Public Property Let RiffVoiceStereoWidth(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 5! Then
        value = 5!
    End If
    rVoices(voiceHandle).StereoWidth = value
End Property

'/**
' * @property RiffVoiceTremoloRate
' * @brief Speed of the volume oscillation LFO in Hz.
' */
Public Property Get RiffVoiceTremoloRate(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceTremoloRate = rVoices(voiceHandle).TremoloRate
End Property
Public Property Let RiffVoiceTremoloRate(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 20! Then
        value = 20!
    End If
    rVoices(voiceHandle).TremoloRate = value
End Property

'/**
' * @property RiffVoiceTremoloDepth
' * @brief Intensity of the volume oscillation effect.
' */
Public Property Get RiffVoiceTremoloDepth(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceTremoloDepth = rVoices(voiceHandle).TremoloDepth
End Property
Public Property Let RiffVoiceTremoloDepth(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 1! Then
        value = 1!
    End If
    rVoices(voiceHandle).TremoloDepth = value
End Property

'/**
' * @property RiffVoiceChorusDepth
' * @brief Wet mix amount for the multi-voice Chorus effect.
' */
Public Property Get RiffVoiceChorusDepth(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceChorusDepth = rVoices(voiceHandle).ChorusDepth
End Property
Public Property Let RiffVoiceChorusDepth(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 1! Then
        value = 1!
    End If
    rVoices(voiceHandle).ChorusDepth = value
End Property

'/**
' * @property RiffVoiceChorusRate
' * @brief LFO rate governing the Chorus pitch modulation.
' */
Public Property Get RiffVoiceChorusRate(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceChorusRate = rVoices(voiceHandle).ChorusRate
End Property
Public Property Let RiffVoiceChorusRate(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0.1! Then
        value = 0.1!
    End If
    If value > 10! Then
        value = 10!
    End If
    rVoices(voiceHandle).ChorusRate = value
End Property

'/**
' * @property RiffVoiceReverbMix
' * @brief Blend of the simulated spatial room reverberation.
' */
Public Property Get RiffVoiceReverbMix(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceReverbMix = rVoices(voiceHandle).ReverbMix
End Property
Public Property Let RiffVoiceReverbMix(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 1! Then
        value = 1!
    End If
    rVoices(voiceHandle).ReverbMix = value
End Property

'/**
' * @property RiffVoiceReverbTime
' * @brief Determines the decay length / simulated room size.
' */
Public Property Get RiffVoiceReverbTime(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceReverbTime = rVoices(voiceHandle).ReverbTime
End Property
Public Property Let RiffVoiceReverbTime(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 0.95! Then
        value = 0.95!
    End If
    rVoices(voiceHandle).ReverbTime = value
End Property

'/**
' * @property RiffVoiceDelayTime
' * @brief Interval between consecutive delay echoes.
' */
Public Property Get RiffVoiceDelayTime(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceDelayTime = rVoices(voiceHandle).DelayTime
End Property
Public Property Let RiffVoiceDelayTime(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 1! Then
        value = 1!
    End If
    rVoices(voiceHandle).DelayTime = value
End Property

'/**
' * @property RiffVoiceDelayFeedback
' * @brief The amount of signal fed back into the delay line to create decaying echoes.
' */
Public Property Get RiffVoiceDelayFeedback(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceDelayFeedback = rVoices(voiceHandle).DelayFeedback
End Property
Public Property Let RiffVoiceDelayFeedback(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 0.95! Then
        value = 0.95!
    End If
    rVoices(voiceHandle).DelayFeedback = value
End Property

'/**
' * @property RiffVoiceDelayMix
' * @brief Blend of the Echo/Delay effect into the main output.
' */
Public Property Get RiffVoiceDelayMix(ByVal voiceHandle As Long) As Single
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    RiffVoiceDelayMix = rVoices(voiceHandle).DelayMix
End Property
Public Property Let RiffVoiceDelayMix(ByVal voiceHandle As Long, ByVal value As Single)
    If Not rCtx.Initialized Or voiceHandle < 0 Or voiceHandle > 31 Then
        Exit Property
    End If
    If value < 0! Then
        value = 0!
    End If
    If value > 1! Then
        value = 1!
    End If
    rVoices(voiceHandle).DelayMix = value
End Property

' ========================================================================================
' PRIVATE ENGINE CORE: HIGH-PERFORMANCE ARRAY CHUNKING CALLBACK
' ========================================================================================

'/**
' * @function RiffTimerCallback
' * @brief Core multimedia timer callback. Executes the DSP pipeline and writes to WASAPI.
' *        Highly optimized to prevent zero-latency buffer underruns.
' * @param hWnd Window Handle (not used).
' * @param uMsg System Message (not used).
' * @param idEvent Timer Identifier (not used).
' * @param dwTime System Time (not used).
' */
#If VBA7 Then
Private Sub RiffTimerCallback(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Private Sub RiffTimerCallback(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    If rCtx.MagicCookie <> &H52494646 Then
        Exit Sub
    End If
    
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
    
    #If VBA7 Then
        Dim pData As LongPtr
        Dim ptr As LongPtr
    #Else
        Dim pData As Long
        Dim ptr As Long
    #End If
    
    hr = vCall(rCtx.AudioClient, 6, VarPtr(padding))
    If hr <> 0 Then
        Exit Sub
    End If
    
    framesAvailable = rCtx.BufferSize - padding
    If framesAvailable <= 0 Then
        Exit Sub
    End If
    
    hr = vCall(rCtx.RenderClient, 3, framesAvailable, VarPtr(pData))
    If hr <> 0 Or pData = 0 Then
        Exit Sub
    End If
    
    RtlMoveMemory VarPtr(nChannels), ByVal (rCtx.MixFormatPtr + 2), 2
    RtlMoveMemory VarPtr(nBlockAlign), ByVal (rCtx.MixFormatPtr + 12), 2
    RtlMoveMemory VarPtr(wBits), ByVal (rCtx.MixFormatPtr + 14), 2
    
    bytesToWrite = framesAvailable * CLng(nBlockAlign)
    align = CLng(nBlockAlign)
    
    rCtx.MasterPeakL = rCtx.MasterPeakL * 0.9!
    rCtx.MasterPeakR = rCtx.MasterPeakR * 0.9!
    
    Dim currentMasterPeakL As Single
    Dim currentMasterPeakR As Single
    currentMasterPeakL = 0!
    currentMasterPeakR = 0!
    
    If wBits = 32 Then
        Dim mixArr32() As Single
        ReDim mixArr32(0 To (bytesToWrite \ 4) - 1)
        
        For i = 0 To 31
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
                    
                    Dim srcArr32() As Single
                    framesNeeded = Int(framesAvailable * ptch) + 2
                    bytesNeeded = framesNeeded * align
                    
                    If Not rVoices(i).IsOscillator Then
                        
                        ptr = rCtx.Buffers(rVoices(i).BufferIndex).BufferPtr
                        readPos = (CLng(pos) \ align) * align
                        
                        If readPos < 0 Then
                            readPos = 0
                        End If
                        
                        bytesAvail = CLng(loopEnd) - readPos
                        If bytesAvail < 0 Then
                            bytesAvail = 0
                        End If
                        
                        ReDim srcArr32(0 To (bytesNeeded \ 4) - 1)
                        
                        If bytesNeeded <= bytesAvail Then
                            RtlMoveMemoryToSingle srcArr32(0), ByVal (ptr + readPos), bytesNeeded
                        Else
                            If bytesAvail > 0 Then
                                RtlMoveMemoryToSingle srcArr32(0), ByVal (ptr + readPos), bytesAvail
                            End If
                            If loopSnd Then
                                remBytes = bytesNeeded - bytesAvail
                                If remBytes > 0 Then
                                    RtlMoveMemoryToSingle srcArr32(bytesAvail \ 4), ByVal (ptr + CLng(loopStart)), remBytes
                                End If
                            End If
                        End If
                        
                    End If
                    
                    dist = rVoices(i).Distortion
                    lp = rVoices(i).LowPass
                    hp = rVoices(i).HighPass
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
                    rmStep = (PI2 * rmFreq) / CDbl(rCtx.SampleRate)
                    
                    trmRate = rVoices(i).TremoloRate
                    trmDepth = rVoices(i).TremoloDepth
                    trmPhase = rVoices(i).TremoloPhase
                    trmStep = (PI2 * trmRate) / CDbl(rCtx.SampleRate)
                    
                    apRate = rVoices(i).AutoPanRate
                    apDepth = rVoices(i).AutoPanDepth
                    apPhase = rVoices(i).AutoPanPhase
                    apStep = (PI2 * apRate) / CDbl(rCtx.SampleRate)
                    
                    cRate = rVoices(i).ChorusRate
                    cDepth = rVoices(i).ChorusDepth
                    cPhase = rVoices(i).ChorusPhase
                    cStep = (PI2 * cRate) / CDbl(rCtx.SampleRate)
                    
                    flgRate = rVoices(i).FlangerRate
                    flgDepth = rVoices(i).FlangerDepth
                    flgFB = rVoices(i).FlangerFeedback
                    flgPhase = rVoices(i).FlangerPhase
                    flgStep = (PI2 * flgRate) / CDbl(rCtx.SampleRate)
                    
                    rMix = rVoices(i).ReverbMix
                    rTime = rVoices(i).ReverbTime
                    rt1 = rVoices(i).RevTap1
                    rt2 = rVoices(i).RevTap2
                    rt3 = rVoices(i).RevTap3
                    rt4 = rVoices(i).RevTap4
                    
                    dTime = rVoices(i).DelayTime
                    dFB = rVoices(i).DelayFeedback
                    dMix = rVoices(i).DelayMix
                    dSamples = (Int(dTime * rCtx.SampleRate) \ 2) * 2
                    dWrite = rVoices(i).RingWritePos
                    dBase = i * 192000
                    
                    fadeState = rVoices(i).fadeState
                    fadeCur = rVoices(i).FadeFramesCurrent
                    fadeTot = rVoices(i).FadeFramesTotal
                    
                    srcIdx = 0
                    writeIdx = 0
                    ptchAlign = ptch * CDbl(align)
                    
                    Dim currentVoicePeakL As Single
                    Dim currentVoicePeakR As Single
                    currentVoicePeakL = 0!
                    currentVoicePeakR = 0!
                    rVoices(i).PeakL = rVoices(i).PeakL * 0.9!
                    rVoices(i).PeakR = rVoices(i).PeakR * 0.9!
                    
                    If nChannels = 2 Then
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
                                Dim oscStep As Double
                                oscStep = (PI2 * rVoices(i).OscFreq) / CDbl(rCtx.SampleRate)
                                
                                Select Case rVoices(i).OscType
                                    Case 0
                                        fL = Sin(rVoices(i).OscPhase)
                                    Case 1
                                        If Sin(rVoices(i).OscPhase) >= 0 Then
                                            fL = 0.5!
                                        Else
                                            fL = -0.5!
                                        End If
                                    Case 2
                                        fL = 2! * (rVoices(i).OscPhase / PI2) - 1!
                                    Case 3
                                        fL = (Rnd() * 2!) - 1!
                                End Select
                                fR = fL
                                
                                rVoices(i).OscPhase = rVoices(i).OscPhase + oscStep
                                If rVoices(i).OscPhase >= PI2 Then
                                    rVoices(i).OscPhase = rVoices(i).OscPhase - PI2
                                End If
                            Else
                                Dim sID As Long
                                sID = Int(srcIdx) * 2
                                If sID + 1 <= UBound(srcArr32) Then
                                    fL = srcArr32(sID)
                                    fR = srcArr32(sID + 1)
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
                                fL = Int(fL * bdSteps) / bdSteps
                                fR = Int(fR * bdSteps) / bdSteps
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
                            
                            rVoices(i).FilterStateL = rVoices(i).FilterStateL + lp * (fL - rVoices(i).FilterStateL)
                            fL = rVoices(i).FilterStateL
                            rVoices(i).FilterStateR = rVoices(i).FilterStateR + lp * (fR - rVoices(i).FilterStateR)
                            fR = rVoices(i).FilterStateR
                            
                            If hp > 0! Then
                                rVoices(i).FilterStateHP_L = rVoices(i).FilterStateHP_L + hp * (fL - rVoices(i).FilterStateHP_L)
                                fL = fL - rVoices(i).FilterStateHP_L
                                rVoices(i).FilterStateHP_R = rVoices(i).FilterStateHP_R + hp * (fR - rVoices(i).FilterStateHP_R)
                                fR = fR - rVoices(i).FilterStateHP_R
                            End If
                            
                            If eqB <> 1! Or eqM <> 1! Or eqT <> 1! Then
                                rVoices(i).EqStateLowL = rVoices(i).EqStateLowL + 0.05! * (fL - rVoices(i).EqStateLowL)
                                rVoices(i).EqStateHighL = rVoices(i).EqStateHighL + 0.4! * (fL - rVoices(i).EqStateHighL)
                                rVoices(i).EqStateLowR = rVoices(i).EqStateLowR + 0.05! * (fR - rVoices(i).EqStateLowR)
                                rVoices(i).EqStateHighR = rVoices(i).EqStateHighR + 0.4! * (fR - rVoices(i).EqStateHighR)
                                
                                Dim midL As Single
                                Dim midR As Single
                                Dim hiL As Single
                                Dim hiR As Single
                                
                                midL = rVoices(i).EqStateHighL - rVoices(i).EqStateLowL
                                midR = rVoices(i).EqStateHighR - rVoices(i).EqStateLowR
                                hiL = fL - rVoices(i).EqStateHighL
                                hiR = fR - rVoices(i).EqStateHighR
                                
                                fL = (rVoices(i).EqStateLowL * eqB) + (midL * eqM) + (hiL * eqT)
                                fR = (rVoices(i).EqStateLowR * eqB) + (midR * eqM) + (hiR * eqT)
                            End If
                            
                            If rmMix > 0! Then
                                Dim rmOsc As Single
                                rmOsc = Sin(rmPhase)
                                fL = fL * (1! - rmMix) + (fL * rmOsc) * rmMix
                                fR = fR * (1! - rmMix) + (fR * rmOsc) * rmMix
                                rmPhase = rmPhase + rmStep
                                If rmPhase > PI2 Then
                                    rmPhase = rmPhase - PI2
                                End If
                            End If
                            
                            If trmDepth > 0! Then
                                Dim trmMult As Single
                                trmMult = 1! - trmDepth * (0.5! + 0.5! * CSng(Sin(trmPhase)))
                                fL = fL * trmMult
                                fR = fR * trmMult
                                trmPhase = trmPhase + trmStep
                                If trmPhase > PI2 Then
                                    trmPhase = trmPhase - PI2
                                End If
                            End If
                            
                            If sWidth <> 1! Then
                                Dim midS As Single
                                Dim sideS As Single
                                midS = (fL + fR) * 0.5!
                                sideS = (fL - fR) * 0.5!
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
                                
                                fDel = (Int((0.002! + 0.005! * CSng(Sin(flgPhase))) * rCtx.SampleRate) \ 2) * 2
                                fRd = dWrite - fDel
                                If fRd < 0 Then
                                    fRd = fRd + 192000
                                End If
                                
                                flgL = rRingBuf(dBase + fRd)
                                flgR = rRingBuf(dBase + fRd + 1)
                                
                                fL = fL + flgL * flgDepth
                                fR = fR + flgR * flgDepth
                                bufInL = bufInL + flgL * flgFB
                                bufInR = bufInR + flgR * flgFB
                                
                                flgPhase = flgPhase + flgStep
                                If flgPhase > PI2 Then
                                    flgPhase = flgPhase - PI2
                                End If
                            End If
                            
                            If cDepth > 0! Then
                                Dim cDelay As Long
                                Dim cRead As Long
                                
                                cDelay = (Int((0.02! + 0.005! * CSng(Sin(cPhase))) * rCtx.SampleRate) \ 2) * 2
                                cRead = dWrite - cDelay
                                If cRead < 0 Then
                                    cRead = cRead + 192000
                                End If
                                
                                fL = fL * (1! - cDepth * 0.5!) + rRingBuf(dBase + cRead) * cDepth
                                fR = fR * (1! - cDepth * 0.5!) + rRingBuf(dBase + cRead + 1) * cDepth
                                cPhase = cPhase + cStep
                                If cPhase > PI2 Then
                                    cPhase = cPhase - PI2
                                End If
                            End If
                            
                            If dMix > 0! And dSamples > 0 Then
                                Dim dRead As Long
                                Dim dL As Single
                                Dim dR As Single
                                
                                dRead = dWrite - dSamples
                                If dRead < 0 Then
                                    dRead = dRead + 192000
                                End If
                                
                                dL = rRingBuf(dBase + dRead)
                                dR = rRingBuf(dBase + dRead + 1)
                                fL = fL + dL * dMix
                                fR = fR + dR * dMix
                                bufInL = bufInL + dL * dFB
                                bufInR = bufInR + dR * dFB
                            End If
                            
                            If rMix > 0! Then
                                Dim r1 As Long
                                Dim r2 As Long
                                Dim r3 As Long
                                Dim r4 As Long
                                Dim revL As Single
                                Dim revR As Single
                                
                                r1 = dWrite - rt1
                                If r1 < 0 Then
                                    r1 = r1 + 192000
                                End If
                                
                                r2 = dWrite - rt2
                                If r2 < 0 Then
                                    r2 = r2 + 192000
                                End If
                                
                                r3 = dWrite - rt3
                                If r3 < 0 Then
                                    r3 = r3 + 192000
                                End If
                                
                                r4 = dWrite - rt4
                                If r4 < 0 Then
                                    r4 = r4 + 192000
                                End If
                                
                                revL = (rRingBuf(dBase + r1) + rRingBuf(dBase + r2) + rRingBuf(dBase + r3) + rRingBuf(dBase + r4)) * 0.25!
                                revR = (rRingBuf(dBase + r1 + 1) + rRingBuf(dBase + r2 + 1) + rRingBuf(dBase + r3 + 1) + rRingBuf(dBase + r4 + 1)) * 0.25!
                                
                                fL = fL + revL * rMix
                                fR = fR + revR * rMix
                                bufInL = bufInL + revL * rTime
                                bufInR = bufInR + revR * rTime
                            End If
                            
                            rRingBuf(dBase + dWrite) = bufInL
                            rRingBuf(dBase + dWrite + 1) = bufInR
                            dWrite = dWrite + 2
                            If dWrite >= 192000 Then
                                dWrite = 0
                            End If
                            
                            If cmpRatio > 1! Then
                                Dim pkL As Single
                                Dim pkR As Single
                                Dim maxPk As Single
                                
                                pkL = Abs(fL)
                                pkR = Abs(fR)
                                maxPk = pkL
                                If pkR > maxPk Then
                                    maxPk = pkR
                                End If
                                
                                If maxPk > cmpEnv Then
                                    cmpEnv = cmpEnv + 0.01! * (maxPk - cmpEnv)
                                Else
                                    cmpEnv = cmpEnv + 0.001! * (maxPk - cmpEnv)
                                End If
                                
                                If cmpEnv > cmpThresh And cmpEnv > 0.0001! Then
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
                                If apPhase > PI2 Then
                                    apPhase = apPhase - PI2
                                End If
                            End If
                            
                            Dim busVol As Single
                            busVol = rCtx.Buses(rVoices(i).busID)
                            
                            vL = rVoices(i).Volume * rCtx.MasterVolume * busVol
                            vR = rVoices(i).Volume * rCtx.MasterVolume * busVol
                            
                            If curPan > 0! Then
                                vL = vL * (1! - curPan)
                            End If
                            If curPan < 0! Then
                                vR = vR * (1! + curPan)
                            End If
                            
                            fL = fL * vL * fadeMult
                            fR = fR * vR * fadeMult
                            
                            If Abs(fL) > currentVoicePeakL Then
                                currentVoicePeakL = Abs(fL)
                            End If
                            If Abs(fR) > currentVoicePeakR Then
                                currentVoicePeakR = Abs(fR)
                            End If
                            
                            mixArr32(writeIdx) = mixArr32(writeIdx) + fL
                            mixArr32(writeIdx + 1) = mixArr32(writeIdx + 1) + fR
                            
                            If Abs(mixArr32(writeIdx)) > currentMasterPeakL Then
                                currentMasterPeakL = Abs(mixArr32(writeIdx))
                            End If
                            If Abs(mixArr32(writeIdx + 1)) > currentMasterPeakR Then
                                currentMasterPeakR = Abs(mixArr32(writeIdx + 1))
                            End If
                            
                            writeIdx = writeIdx + 2
                            srcIdx = srcIdx + ptch
                            pos = pos + ptchAlign
                        Next frame
                    End If
                    
                    If currentVoicePeakL > rVoices(i).PeakL Then
                        rVoices(i).PeakL = currentVoicePeakL
                    End If
                    If currentVoicePeakR > rVoices(i).PeakR Then
                        rVoices(i).PeakR = currentVoicePeakR
                    End If
                    
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
        
        If currentMasterPeakL > rCtx.MasterPeakL Then
            rCtx.MasterPeakL = currentMasterPeakL
        End If
        If currentMasterPeakR > rCtx.MasterPeakR Then
            rCtx.MasterPeakR = currentMasterPeakR
        End If
        
        RtlMoveMemory ByVal pData, VarPtr(mixArr32(0)), bytesToWrite

    ' ------------------------------------------------------------------------------------
    ' FALLBACK PATH: 16-Bit Integer (Legacy Audio Hardware)
    ' ------------------------------------------------------------------------------------
    ElseIf wBits = 16 Then
        Dim mixArr16() As Integer
        ReDim mixArr16(0 To (bytesToWrite \ 2) - 1)
        
        For i = 0 To 31
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
                    
                    Dim srcArr16() As Integer
                    framesNeeded = Int(framesAvailable * ptch) + 2
                    bytesNeeded = framesNeeded * align
                    
                    If Not rVoices(i).IsOscillator Then
                        
                        ptr = rCtx.Buffers(rVoices(i).BufferIndex).BufferPtr
                        readPos = (CLng(pos) \ align) * align
                        
                        If readPos < 0 Then
                            readPos = 0
                        End If
                        
                        bytesAvail = CLng(loopEnd) - readPos
                        If bytesAvail < 0 Then
                            bytesAvail = 0
                        End If
                        
                        ReDim srcArr16(0 To (bytesNeeded \ 2) - 1)
                        
                        If bytesNeeded <= bytesAvail Then
                            RtlMoveMemoryToInteger srcArr16(0), ByVal (ptr + readPos), bytesNeeded
                        Else
                            If bytesAvail > 0 Then
                                RtlMoveMemoryToInteger srcArr16(0), ByVal (ptr + readPos), bytesAvail
                            End If
                            If loopSnd Then
                                remBytes = bytesNeeded - bytesAvail
                                If remBytes > 0 Then
                                    RtlMoveMemoryToInteger srcArr16(bytesAvail \ 2), ByVal (ptr + CLng(loopStart)), remBytes
                                End If
                            End If
                        End If
                        
                    End If
                    
                    dist = rVoices(i).Distortion
                    lp = rVoices(i).LowPass
                    hp = rVoices(i).HighPass
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
                    rmStep = (PI2 * rmFreq) / CDbl(rCtx.SampleRate)
                    
                    trmRate = rVoices(i).TremoloRate
                    trmDepth = rVoices(i).TremoloDepth
                    trmPhase = rVoices(i).TremoloPhase
                    trmStep = (PI2 * trmRate) / CDbl(rCtx.SampleRate)
                    
                    apRate = rVoices(i).AutoPanRate
                    apDepth = rVoices(i).AutoPanDepth
                    apPhase = rVoices(i).AutoPanPhase
                    apStep = (PI2 * apRate) / CDbl(rCtx.SampleRate)
                    
                    cRate = rVoices(i).ChorusRate
                    cDepth = rVoices(i).ChorusDepth
                    cPhase = rVoices(i).ChorusPhase
                    cStep = (PI2 * cRate) / CDbl(rCtx.SampleRate)
                    
                    flgRate = rVoices(i).FlangerRate
                    flgDepth = rVoices(i).FlangerDepth
                    flgFB = rVoices(i).FlangerFeedback
                    flgPhase = rVoices(i).FlangerPhase
                    flgStep = (PI2 * flgRate) / CDbl(rCtx.SampleRate)
                    
                    rMix = rVoices(i).ReverbMix
                    rTime = rVoices(i).ReverbTime
                    rt1 = rVoices(i).RevTap1
                    rt2 = rVoices(i).RevTap2
                    rt3 = rVoices(i).RevTap3
                    rt4 = rVoices(i).RevTap4
                    
                    dTime = rVoices(i).DelayTime
                    dFB = rVoices(i).DelayFeedback
                    dMix = rVoices(i).DelayMix
                    dSamples = (Int(dTime * rCtx.SampleRate) \ 2) * 2
                    dWrite = rVoices(i).RingWritePos
                    dBase = i * 192000
                    
                    fadeState = rVoices(i).fadeState
                    fadeCur = rVoices(i).FadeFramesCurrent
                    fadeTot = rVoices(i).FadeFramesTotal
                    
                    srcIdx = 0
                    writeIdx = 0
                    ptchAlign = ptch * CDbl(align)
                    
                    Dim cVoicePeakL16 As Single
                    Dim cVoicePeakR16 As Single
                    cVoicePeakL16 = 0!
                    cVoicePeakR16 = 0!
                    rVoices(i).PeakL = rVoices(i).PeakL * 0.9!
                    rVoices(i).PeakR = rVoices(i).PeakR * 0.9!
                    
                    If nChannels = 2 Then
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
                                Dim oscStep16 As Double
                                oscStep16 = (PI2 * rVoices(i).OscFreq) / CDbl(rCtx.SampleRate)
                                
                                Select Case rVoices(i).OscType
                                    Case 0
                                        fL = Sin(rVoices(i).OscPhase)
                                    Case 1
                                        If Sin(rVoices(i).OscPhase) >= 0 Then
                                            fL = 0.5!
                                        Else
                                            fL = -0.5!
                                        End If
                                    Case 2
                                        fL = 2! * (rVoices(i).OscPhase / PI2) - 1!
                                    Case 3
                                        fL = (Rnd() * 2!) - 1!
                                End Select
                                fR = fL
                                
                                rVoices(i).OscPhase = rVoices(i).OscPhase + oscStep16
                                If rVoices(i).OscPhase >= PI2 Then
                                    rVoices(i).OscPhase = rVoices(i).OscPhase - PI2
                                End If
                            Else
                                Dim sID16 As Long
                                sID16 = Int(srcIdx) * 2
                                If sID16 + 1 <= UBound(srcArr16) Then
                                    fL = CSng(srcArr16(sID16)) * 3.051758E-05!
                                    fR = CSng(srcArr16(sID16 + 1)) * 3.051758E-05!
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
                                fL = Int(fL * bdSteps) / bdSteps
                                fR = Int(fR * bdSteps) / bdSteps
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
                            
                            rVoices(i).FilterStateL = rVoices(i).FilterStateL + lp * (fL - rVoices(i).FilterStateL)
                            fL = rVoices(i).FilterStateL
                            rVoices(i).FilterStateR = rVoices(i).FilterStateR + lp * (fR - rVoices(i).FilterStateR)
                            fR = rVoices(i).FilterStateR
                            
                            If hp > 0! Then
                                rVoices(i).FilterStateHP_L = rVoices(i).FilterStateHP_L + hp * (fL - rVoices(i).FilterStateHP_L)
                                fL = fL - rVoices(i).FilterStateHP_L
                                rVoices(i).FilterStateHP_R = rVoices(i).FilterStateHP_R + hp * (fR - rVoices(i).FilterStateHP_R)
                                fR = fR - rVoices(i).FilterStateHP_R
                            End If
                            
                            If eqB <> 1! Or eqM <> 1! Or eqT <> 1! Then
                                rVoices(i).EqStateLowL = rVoices(i).EqStateLowL + 0.05! * (fL - rVoices(i).EqStateLowL)
                                rVoices(i).EqStateHighL = rVoices(i).EqStateHighL + 0.4! * (fL - rVoices(i).EqStateHighL)
                                rVoices(i).EqStateLowR = rVoices(i).EqStateLowR + 0.05! * (fR - rVoices(i).EqStateLowR)
                                rVoices(i).EqStateHighR = rVoices(i).EqStateHighR + 0.4! * (fR - rVoices(i).EqStateHighR)
                                
                                Dim mL16 As Single
                                Dim mR16 As Single
                                Dim hL16 As Single
                                Dim hR16 As Single
                                
                                mL16 = rVoices(i).EqStateHighL - rVoices(i).EqStateLowL
                                mR16 = rVoices(i).EqStateHighR - rVoices(i).EqStateLowR
                                hL16 = fL - rVoices(i).EqStateHighL
                                hR16 = fR - rVoices(i).EqStateHighR
                                
                                fL = (rVoices(i).EqStateLowL * eqB) + (mL16 * eqM) + (hL16 * eqT)
                                fR = (rVoices(i).EqStateLowR * eqB) + (mR16 * eqM) + (hR16 * eqT)
                            End If
                            
                            If rmMix > 0! Then
                                Dim rmOsc16 As Single
                                rmOsc16 = Sin(rmPhase)
                                fL = fL * (1! - rmMix) + (fL * rmOsc16) * rmMix
                                fR = fR * (1! - rmMix) + (fR * rmOsc16) * rmMix
                                rmPhase = rmPhase + rmStep
                                If rmPhase > PI2 Then
                                    rmPhase = rmPhase - PI2
                                End If
                            End If
                            
                            If trmDepth > 0! Then
                                Dim trmM16 As Single
                                trmM16 = 1! - trmDepth * (0.5! + 0.5! * CSng(Sin(trmPhase)))
                                fL = fL * trmM16
                                fR = fR * trmM16
                                trmPhase = trmPhase + trmStep
                                If trmPhase > PI2 Then
                                    trmPhase = trmPhase - PI2
                                End If
                            End If
                            
                            If sWidth <> 1! Then
                                Dim m16 As Single
                                Dim sd16 As Single
                                m16 = (fL + fR) * 0.5!
                                sd16 = (fL - fR) * 0.5!
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
                                
                                fDel16 = (Int((0.002! + 0.005! * CSng(Sin(flgPhase))) * rCtx.SampleRate) \ 2) * 2
                                fRd16 = dWrite - fDel16
                                If fRd16 < 0 Then
                                    fRd16 = fRd16 + 192000
                                End If
                                
                                flgL16 = rRingBuf(dBase + fRd16)
                                flgR16 = rRingBuf(dBase + fRd16 + 1)
                                
                                fL = fL + flgL16 * flgDepth
                                fR = fR + flgR16 * flgDepth
                                bufInL16 = bufInL16 + flgL16 * flgFB
                                bufInR16 = bufInR16 + flgR16 * flgFB
                                
                                flgPhase = flgPhase + flgStep
                                If flgPhase > PI2 Then
                                    flgPhase = flgPhase - PI2
                                End If
                            End If
                            
                            If cDepth > 0! Then
                                Dim cDel16 As Long
                                Dim cRd16 As Long
                                
                                cDel16 = (Int((0.02! + 0.005! * CSng(Sin(cPhase))) * rCtx.SampleRate) \ 2) * 2
                                cRd16 = dWrite - cDel16
                                If cRd16 < 0 Then
                                    cRd16 = cRd16 + 192000
                                End If
                                
                                fL = fL * (1! - cDepth * 0.5!) + rRingBuf(dBase + cRd16) * cDepth
                                fR = fR * (1! - cDepth * 0.5!) + rRingBuf(dBase + cRd16 + 1) * cDepth
                                cPhase = cPhase + cStep
                                If cPhase > PI2 Then
                                    cPhase = cPhase - PI2
                                End If
                            End If
                            
                            If dMix > 0! And dSamples > 0 Then
                                Dim dR16 As Long
                                Dim dL16 As Single
                                Dim dR16_2 As Single
                                
                                dR16 = dWrite - dSamples
                                If dR16 < 0 Then
                                    dR16 = dR16 + 192000
                                End If
                                
                                dL16 = rRingBuf(dBase + dR16)
                                dR16_2 = rRingBuf(dBase + dR16 + 1)
                                fL = fL + dL16 * dMix
                                fR = fR + dR16_2 * dMix
                                bufInL16 = bufInL16 + dL16 * dFB
                                bufInR16 = bufInR16 + dR16_2 * dFB
                            End If
                            
                            If rMix > 0! Then
                                Dim rr1 As Long
                                Dim rr2 As Long
                                Dim rr3 As Long
                                Dim rr4 As Long
                                Dim revL16 As Single
                                Dim revR16 As Single
                                
                                rr1 = dWrite - rt1
                                If rr1 < 0 Then
                                    rr1 = rr1 + 192000
                                End If
                                
                                rr2 = dWrite - rt2
                                If rr2 < 0 Then
                                    rr2 = rr2 + 192000
                                End If
                                
                                rr3 = dWrite - rt3
                                If rr3 < 0 Then
                                    rr3 = rr3 + 192000
                                End If
                                
                                rr4 = dWrite - rt4
                                If rr4 < 0 Then
                                    rr4 = rr4 + 192000
                                End If
                                
                                revL16 = (rRingBuf(dBase + rr1) + rRingBuf(dBase + rr2) + rRingBuf(dBase + rr3) + rRingBuf(dBase + rr4)) * 0.25!
                                revR16 = (rRingBuf(dBase + rr1 + 1) + rRingBuf(dBase + rr2 + 1) + rRingBuf(dBase + rr3 + 1) + rRingBuf(dBase + rr4 + 1)) * 0.25!
                                
                                fL = fL + revL16 * rMix
                                fR = fR + revR16 * rMix
                                bufInL16 = bufInL16 + revL16 * rTime
                                bufInR16 = bufInR16 + revR16 * rTime
                            End If
                            
                            rRingBuf(dBase + dWrite) = bufInL16
                            rRingBuf(dBase + dWrite + 1) = bufInR16
                            dWrite = dWrite + 2
                            If dWrite >= 192000 Then
                                dWrite = 0
                            End If
                            
                            If cmpRatio > 1! Then
                                Dim pkL16 As Single
                                Dim pkR16 As Single
                                Dim maxPk16 As Single
                                
                                pkL16 = Abs(fL)
                                pkR16 = Abs(fR)
                                maxPk16 = pkL16
                                If pkR16 > maxPk16 Then
                                    maxPk16 = pkR16
                                End If
                                
                                If maxPk16 > cmpEnv Then
                                    cmpEnv = cmpEnv + 0.01! * (maxPk16 - cmpEnv)
                                Else
                                    cmpEnv = cmpEnv + 0.001! * (maxPk16 - cmpEnv)
                                End If
                                
                                If cmpEnv > cmpThresh And cmpEnv > 0.0001! Then
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
                                If apPhase > PI2 Then
                                    apPhase = apPhase - PI2
                                End If
                            End If
                            
                            Dim bVol16 As Single
                            bVol16 = rCtx.Buses(rVoices(i).busID)
                            
                            vL = rVoices(i).Volume * rCtx.MasterVolume * bVol16
                            vR = rVoices(i).Volume * rCtx.MasterVolume * bVol16
                            
                            If cPan16 > 0! Then
                                vL = vL * (1! - cPan16)
                            End If
                            If cPan16 < 0! Then
                                vR = vR * (1! + cPan16)
                            End If
                            
                            fL = fL * vL * fadeMult
                            fR = fR * vR * fadeMult
                            
                            If Abs(fL) > cVoicePeakL16 Then
                                cVoicePeakL16 = Abs(fL)
                            End If
                            If Abs(fR) > cVoicePeakR16 Then
                                cVoicePeakR16 = Abs(fR)
                            End If
                            
                            l1 = CLng(mixArr16(writeIdx)) + CLng(fL * 32767!)
                            l2 = CLng(mixArr16(writeIdx + 1)) + CLng(fR * 32767!)
                            
                            If l1 > 32767 Then
                                l1 = 32767
                            ElseIf l1 < -32768 Then
                                l1 = -32768
                            End If
                            
                            If l2 > 32767 Then
                                l2 = 32767
                            ElseIf l2 < -32768 Then
                                l2 = -32768
                            End If
                            
                            mixArr16(writeIdx) = CInt(l1)
                            mixArr16(writeIdx + 1) = CInt(l2)
                            
                            If Abs(fL) > currentMasterPeakL Then
                                currentMasterPeakL = Abs(fL)
                            End If
                            If Abs(fR) > currentMasterPeakR Then
                                currentMasterPeakR = Abs(fR)
                            End If
                            
                            writeIdx = writeIdx + 2
                            srcIdx = srcIdx + ptch
                            pos = pos + ptchAlign
                        Next frame
                    End If
                    
                    If cVoicePeakL16 > rVoices(i).PeakL Then
                        rVoices(i).PeakL = cVoicePeakL16
                    End If
                    If cVoicePeakR16 > rVoices(i).PeakR Then
                        rVoices(i).PeakR = cVoicePeakR16
                    End If
                    
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
        
        If currentMasterPeakL > rCtx.MasterPeakL Then
            rCtx.MasterPeakL = currentMasterPeakL
        End If
        If currentMasterPeakR > rCtx.MasterPeakR Then
            rCtx.MasterPeakR = currentMasterPeakR
        End If
        
        RtlMoveMemory ByVal pData, VarPtr(mixArr16(0)), bytesToWrite
    End If
    
    vCall rCtx.RenderClient, 4, framesAvailable, 0&
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
        Const THUNK_SIZE As Long = 1024
    #Else
        Const THUNK_SIZE As Long = 512
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
            pEbMode = GetProcAddressOrdinal(hVbe, 1&)
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
        ReDim opcodes(0 To (Len(hexStr) \ 2) - 1)
        For i = 0 To UBound(opcodes)
            opcodes(i) = CByte("&H" & Mid$(hexStr, (i * 2) + 1, 2))
        Next i
        RtlMoveMemory VarPtr(opcodes(26)), VarPtr(pEbMode), 8
        RtlMoveMemory VarPtr(opcodes(63)), VarPtr(pKill), 8
        RtlMoveMemory VarPtr(opcodes(102)), VarPtr(pCallback), 8
    #Else
        hexStr = "5589E5B80000000085C07421FFD083F801741A83F80274158B4510508B450850B8" & _
                 "0000000085C0741FFFD0EB1B8B4514508B4510508B450C508B450850B800000000" & _
                 "85C07402FFD05DC21000"
        ReDim opcodes(0 To (Len(hexStr) \ 2) - 1)
        For i = 0 To UBound(opcodes)
            opcodes(i) = CByte("&H" & Mid$(hexStr, (i * 2) + 1, 2))
        Next i
        RtlMoveMemory VarPtr(opcodes(4)), VarPtr(pEbMode), 4
        RtlMoveMemory VarPtr(opcodes(33)), VarPtr(pKill), 4
        RtlMoveMemory VarPtr(opcodes(62)), VarPtr(pCallback), 4
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
    
    #If Win64 Then
        Dim pNullPtr As LongLong
        Dim hnsDur As LongLong
        Dim hnsPer As LongLong
        pNullPtr = CLngLng(0)
        hnsDur = CLngLng(1500000)
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
        hnsDur = CCur(150)
        hnsPer = CCur(0)
    #End If
    
    IIDFromString StrPtr("{BCDE0395-E52F-467C-8E3D-C4579291692E}"), clsidEnum
    IIDFromString StrPtr("{A95664D2-9614-4F35-A746-DE8DB63617E6}"), iidEnum
    IIDFromString StrPtr("{1CB9AD4C-DBFA-4c32-B178-C2F568A703B2}"), iidAudio
    IIDFromString StrPtr("{F294ACFC-3146-4483-A7BF-ADDCA7C260E2}"), iidRender
    
    hr = CoCreateInstance(clsidEnum, 0, CLSCTX_ALL, iidEnum, rCtx.DeviceEnumerator)
    If hr <> 0 Or rCtx.DeviceEnumerator = 0 Then
        Exit Function
    End If
    
    hr = vCall(rCtx.DeviceEnumerator, 4, eRender, eConsole, VarPtr(rCtx.Device))
    If hr <> 0 Or rCtx.Device = 0 Then
        Exit Function
    End If
    
    hr = vCall(rCtx.Device, 3, VarPtr(iidAudio), CLSCTX_ALL, pNullPtr, VarPtr(rCtx.AudioClient))
    If hr <> 0 Or rCtx.AudioClient = 0 Then
        Exit Function
    End If
    
    hr = vCall(rCtx.AudioClient, 8, VarPtr(rCtx.MixFormatPtr))
    If hr <> 0 Or rCtx.MixFormatPtr = 0 Then
        Exit Function
    End If
    
    RtlMoveMemory VarPtr(rCtx.SampleRate), ByVal (rCtx.MixFormatPtr + 4), 4
    RtlMoveMemory VarPtr(rCtx.AvgBytesPerSec), ByVal (rCtx.MixFormatPtr + 8), 4
    
    hr = vCall(rCtx.AudioClient, 3, AUDCLNT_SHAREMODE_SHARED, &H80000000, hnsDur, hnsPer, rCtx.MixFormatPtr, pNullPtr)
    If hr <> 0 Then
        hr = vCall(rCtx.AudioClient, 3, AUDCLNT_SHAREMODE_SHARED, 0&, hnsDur, hnsPer, rCtx.MixFormatPtr, pNullPtr)
    End If
    
    If hr <> 0 Then
        Exit Function
    End If
    
    hr = vCall(rCtx.AudioClient, 4, VarPtr(rCtx.BufferSize))
    If hr <> 0 Then
        Exit Function
    End If
    
    hr = vCall(rCtx.AudioClient, 14, VarPtr(iidRender), VarPtr(rCtx.RenderClient))
    If hr <> 0 Or rCtx.RenderClient = 0 Then
        Exit Function
    End If
    
    hr = vCall(rCtx.AudioClient, 10)
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
        vCall rCtx.RenderClient, 2
        rCtx.RenderClient = 0
    End If
    
    If rCtx.AudioClient <> 0 Then
        vCall rCtx.AudioClient, 11
        vCall rCtx.AudioClient, 2
        rCtx.AudioClient = 0
    End If
    
    If rCtx.Device <> 0 Then
        vCall rCtx.Device, 2
        rCtx.Device = 0
    End If
    
    If rCtx.DeviceEnumerator <> 0 Then
        vCall rCtx.DeviceEnumerator, 2
        rCtx.DeviceEnumerator = 0
    End If
    
    If rCtx.MixFormatPtr <> 0 Then
        CoTaskMemFree rCtx.MixFormatPtr
        rCtx.MixFormatPtr = 0
    End If
End Sub

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
    
    #If Win64 Then
        offset = vTableIndex * 8
    #Else
        offset = vTableIndex * 4
    #End If
    
    argCount = UBound(args) - LBound(args) + 1
    
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
