Option Explicit

' Examples for the Riff.bas audio engine.
' Import this BAS beside Riff.bas, then run Showcase.
' Original code by @C-Johnson-83 but re-adapted to new Riff structure.

Private Sub LWait(ByVal seconds As Single)
    Dim t As Double
    t = Timer
    Do While Timer < t + seconds
        DoEvents
    Loop
End Sub

Private Function LStartEngine() As Boolean
    If RiffIsInitialized Then
        LStartEngine = True
        Exit Function
    End If

    If Not RiffOpen() Then
        MsgBox "Failed to initialize the Riff audio engine." & vbCrLf & _
               "Check that your audio device is connected and working.", _
               vbCritical, "Riff Engine"
        LStartEngine = False
    Else
        LStartEngine = True
    End If
End Function

Private Function LAsk(ByVal title As String, ByVal description As String) As Boolean
    ' No icon flag here. Icon flags cause the Windows message sound.
    LAsk = (MsgBox(description & vbCrLf & vbCrLf & _
                  "OK = run this demo" & vbCrLf & _
                  "Cancel = skip it", vbOKCancel, title) = vbOK)
End Function

Private Sub LStatus(ByVal text As String)
    Application.StatusBar = text
    DoEvents
End Sub

Private Sub LStopClean(Optional ByVal fadeSec As Single = 0.25!)
    ' Fade active voices briefly so the demo does not click/pop when stopping.
    Dim i As Long
    For i = 0 To RiffMaxVoices - 1
        If RiffVoiceIsPlaying(i) Then RiffFadeOut i, fadeSec
    Next i
    LWait fadeSec + 0.05!
    RiffStopAll
    Application.StatusBar = False
End Sub

Private Sub LSetVoice(ByVal v As Long, ByVal vol As Single, ByVal pan As Single)
    If v < 0 Then Exit Sub
    RiffVoiceVolume(v) = vol
    RiffVoicePan(v) = pan
End Sub

Private Sub LPlayMinorPad(ByRef v1 As Long, ByRef v2 As Long, ByRef v3 As Long, _
                          Optional ByVal rootHz As Single = 220!, _
                          Optional ByVal waveType As RiffWaveType = RiffWaveSine)
    ' root, minor third, fifth
    v1 = RiffPlayOscillator(waveType, rootHz)
    v2 = RiffPlayOscillator(waveType, rootHz * 1.189207!)
    v3 = RiffPlayOscillator(waveType, rootHz * 1.498307!)

    LSetVoice v1, 0.12!, -0.35!
    LSetVoice v2, 0.1!, 0!
    LSetVoice v3, 0.1!, 0.35!
End Sub

Private Sub LApplyReverb(ByVal v As Long, ByVal mix As Single, ByVal roomTime As Single)
    If v < 0 Then Exit Sub
    RiffVoiceReverbMix(v) = mix
    RiffVoiceReverbTime(v) = roomTime
End Sub

Private Sub LApplyDelay(ByVal v As Long, ByVal delayTime As Single, ByVal feedback As Single, ByVal mix As Single)
    If v < 0 Then Exit Sub
    RiffVoiceDelayTime(v) = delayTime
    RiffVoiceDelayFeedback(v) = feedback
    RiffVoiceDelayMix(v) = mix
End Sub

Public Sub Showcase()
    If Not LStartEngine() Then Exit Sub

    Dim msg As String
    msg = "RIFF SHOWCASE" & vbCrLf & String(34, "=") & vbCrLf & vbCrLf & _
          "These are examples that make each effect easier to hear." & vbCrLf & vbCrLf & _
          "They avoid the Windows MessageBox ding during normal prompts." & vbCrLf & _
          "Watch the Excel status bar for the current stage."

    If MsgBox(msg & vbCrLf & vbCrLf & "Start?", vbOKCancel, "Riff Showcase") <> vbOK Then Exit Sub

    If LAsk("1. Reverb Room Size", "A sustained chord moves from dry, to small room, to huge cathedral reverb.") Then Demo_ReverbRoomSize
    If LAsk("2. Delay Echo", "A pulsing tone changes from slapback echo to a long spacey echo.") Then Demo_DelayEcho
    If LAsk("3. Chorus vs Flanger", "A pad first gets chorus shimmer, then a flanger jet sweep.") Then Demo_ChorusFlanger
    If LAsk("4. Filter Sweep", "A saw wave opens and closes low-pass and high-pass filters so the tone change is obvious.") Then Demo_FilterSweep
    If LAsk("5. EQ Shapes", "A chord cycles through flat, bass-heavy, bright, and scooped-mid EQ.") Then Demo_EQShapes
    If LAsk("6. Distortion + Bitcrusher", "A clean tone becomes overdriven, then turns into harsh lo-fi crushed audio.") Then Demo_DistortionBitcrusher
    If LAsk("7. Tremolo + AutoPan", "A held tone pulses in volume, then sweeps left and right. Headphones help.") Then Demo_TremoloAutoPan
    If LAsk("8. Ring Mod Metallic", "A clean tone turns bell-like, robotic, then metallic/alien.") Then Demo_RingModMetallic
    If LAsk("9. Pitch + Fade", "A tone fades in, bends down/up in pitch, then fades out.") Then Demo_PitchFade
    If LAsk("10. Presets & Noise", "Instantly apply professional character presets and listen to new organic noise types.") Then Demo_PresetsNoise

    LStopClean 0.2!
    RiffClose
    MsgBox "Showcase complete.", vbOKOnly, "Riff Showcase"
End Sub

Public Sub Demo_ReverbRoomSize()
    If Not LStartEngine() Then Exit Sub

    Dim a As Long, b As Long, c As Long
    LPlayMinorPad a, b, c, 196!, 0

    LStatus "Reverb 1/4 - Dry chord, no reverb"
    LWait 3

    LApplyReverb a, 0.22!, 0.25!: LApplyReverb b, 0.22!, 0.25!: LApplyReverb c, 0.22!, 0.25!
    LStatus "Reverb 2/4 - Small room: short tail"
    LWait 4

    LApplyReverb a, 0.5!, 0.6!: LApplyReverb b, 0.5!, 0.6!: LApplyReverb c, 0.5!, 0.6!
    LStatus "Reverb 3/4 - Hall: longer tail and more wet signal"
    LWait 4

    LApplyReverb a, 0.85!, 0.95!: LApplyReverb b, 0.85!, 0.95!: LApplyReverb c, 0.85!, 0.95!
    LStatus "Reverb 4/4 - Cathedral: huge wash/decay"
    LWait 6

    LStopClean 1!
End Sub

Public Sub Demo_DelayEcho()
    If Not LStartEngine() Then Exit Sub

    Dim v As Long, i As Long
    v = RiffPlayOscillator(RiffWaveSquare, 330!) ' square wave cuts through delay better
    LSetVoice v, 0.1!, 0!

    LStatus "Delay 1/4 - Dry pulsing tone"
    For i = 1 To 8
        RiffVoiceVolume(v) = 0.12!: LWait 0.14!
        RiffVoiceVolume(v) = 0!: LWait 0.22!
    Next i

    RiffVoiceVolume(v) = 0.12!
    LApplyDelay v, 0.11!, 0.25!, 0.45!
    LStatus "Delay 2/4 - Slapback: very short echo"
    For i = 1 To 8
        RiffVoiceVolume(v) = 0.12!: LWait 0.14!
        RiffVoiceVolume(v) = 0!: LWait 0.22!
    Next i

    RiffVoiceVolume(v) = 0.12!
    LApplyDelay v, 0.32!, 0.55!, 0.55!
    LStatus "Delay 3/4 - Rhythmic echo: repeats are easier to hear"
    For i = 1 To 8
        RiffVoiceVolume(v) = 0.12!: LWait 0.14!
        RiffVoiceVolume(v) = 0!: LWait 0.34!
    Next i

    RiffVoiceVolume(v) = 0.12!
    LApplyDelay v, 0.55!, 0.78!, 0.68!
    LStatus "Delay 4/4 - Long feedback echo: trailing repeats"
    For i = 1 To 5
        RiffVoiceVolume(v) = 0.12!: LWait 0.18!
        RiffVoiceVolume(v) = 0!: LWait 0.65!
    Next i

    LWait 3
    LStopClean 0.5!
End Sub

Public Sub Demo_ChorusFlanger()
    If Not LStartEngine() Then Exit Sub

    Dim a As Long, b As Long, c As Long
    LPlayMinorPad a, b, c, 246.94!, 0

    LStatus "Chorus/Flanger 1/3 - Dry pad"
    LWait 4

    RiffVoiceChorusDepth(a) = 0.65!: RiffVoiceChorusRate(a) = 1.25!
    RiffVoiceChorusDepth(b) = 0.65!: RiffVoiceChorusRate(b) = 1.25!
    RiffVoiceChorusDepth(c) = 0.65!: RiffVoiceChorusRate(c) = 1.25!
    LStatus "Chorus/Flanger 2/3 - Chorus: thicker, wider shimmer"
    LWait 6

    RiffVoiceFlangerDepth(a) = 0.8!: RiffVoiceFlangerRate(a) = 0.25!: RiffVoiceFlangerFeedback(a) = 0.65!
    RiffVoiceFlangerDepth(b) = 0.8!: RiffVoiceFlangerRate(b) = 0.25!: RiffVoiceFlangerFeedback(b) = 0.65!
    RiffVoiceFlangerDepth(c) = 0.8!: RiffVoiceFlangerRate(c) = 0.25!: RiffVoiceFlangerFeedback(c) = 0.65!
    LStatus "Chorus/Flanger 3/3 - Flanger: slow jet-like sweep"
    LWait 8

    LStopClean 0.75!
End Sub

Public Sub Demo_FilterSweep()
    If Not LStartEngine() Then Exit Sub

    Dim v As Long, i As Long
    v = RiffPlayOscillator(RiffWaveSawtooth, 110!) ' saw wave has lots of harmonics for filters
    LSetVoice v, 0.12!, 0!

    LStatus "Filter 1/4 - Bright raw saw wave"
    LWait 3

    LStatus "Filter 2/4 - Closing low-pass: sound gets muffled"
    For i = 100 To 5 Step -5
        RiffVoiceLowPass(v) = i / 100!
        LWait 0.18!
    Next i
    LWait 2

    LStatus "Filter 3/4 - Opening low-pass: brightness returns"
    For i = 5 To 100 Step 5
        RiffVoiceLowPass(v) = i / 100!
        LWait 0.18!
    Next i
    RiffVoiceLowPass(v) = 1!
    LWait 1

    LStatus "Filter 4/4 - Raising high-pass: low end disappears, gets thin"
    For i = 0 To 80 Step 4
        RiffVoiceHighPass(v) = i / 100!
        LWait 0.16!
    Next i
    LWait 3

    LStopClean 0.5!
End Sub

Public Sub Demo_EQShapes()
    If Not LStartEngine() Then Exit Sub

    Dim a As Long, b As Long, c As Long
    LPlayMinorPad a, b, c, 164.81!, 2

    LStatus "EQ 1/4 - Flat EQ"
    LWait 4

    RiffVoiceEqBass(a) = 2.8!: RiffVoiceEqBass(b) = 2.8!: RiffVoiceEqBass(c) = 2.8!
    RiffVoiceEqMid(a) = 0.75!: RiffVoiceEqMid(b) = 0.75!: RiffVoiceEqMid(c) = 0.75!
    RiffVoiceEqTreble(a) = 0.55!: RiffVoiceEqTreble(b) = 0.55!: RiffVoiceEqTreble(c) = 0.55!
    LStatus "EQ 2/4 - Bass heavy: thicker and darker"
    LWait 5

    RiffVoiceEqBass(a) = 0.45!: RiffVoiceEqBass(b) = 0.45!: RiffVoiceEqBass(c) = 0.45!
    RiffVoiceEqMid(a) = 0.9!: RiffVoiceEqMid(b) = 0.9!: RiffVoiceEqMid(c) = 0.9!
    RiffVoiceEqTreble(a) = 3!: RiffVoiceEqTreble(b) = 3!: RiffVoiceEqTreble(c) = 3!
    LStatus "EQ 3/4 - Bright: more edge and sparkle"
    LWait 5

    RiffVoiceEqBass(a) = 2.2!: RiffVoiceEqBass(b) = 2.2!: RiffVoiceEqBass(c) = 2.2!
    RiffVoiceEqMid(a) = 0.25!: RiffVoiceEqMid(b) = 0.25!: RiffVoiceEqMid(c) = 0.25!
    RiffVoiceEqTreble(a) = 2.2!: RiffVoiceEqTreble(b) = 2.2!: RiffVoiceEqTreble(c) = 2.2!
    LStatus "EQ 4/4 - Scooped mids: hollow V-shape"
    LWait 5

    LStopClean 0.5!
End Sub

Public Sub Demo_DistortionBitcrusher()
    If Not LStartEngine() Then Exit Sub

    Dim v As Long
    v = RiffPlayOscillator(RiffWaveSine, 146.83!)
    LSetVoice v, 0.12!, 0!

    LStatus "Distortion/Crusher 1/5 - Clean sine wave"
    LWait 3

    RiffVoiceDistortion(v) = 2.5!
    LStatus "Distortion/Crusher 2/5 - Mild overdrive: warmer/rougher"
    LWait 4

    RiffVoiceDistortion(v) = 10!
    LStatus "Distortion/Crusher 3/5 - Heavy distortion: fuzzy/angry"
    LWait 4

    RiffVoiceBitDepth(v) = 8!
    LStatus "Distortion/Crusher 4/5 - 8-bit crush: gritty steps"
    LWait 4

    RiffVoiceBitDepth(v) = 4!
    RiffVoiceSampleRateReduction(v) = 10
    LStatus "Distortion/Crusher 5/5 - Extreme crush: broken/robotic aliasing"
    LWait 5

    LStopClean 0.5!
End Sub

Public Sub Demo_TremoloAutoPan()
    If Not LStartEngine() Then Exit Sub

    Dim v As Long
    v = RiffPlayOscillator(RiffWaveSine, 261.63!)
    LSetVoice v, 0.13!, 0!

    LStatus "Tremolo/AutoPan 1/4 - Plain centered tone"
    LWait 3

    RiffVoiceTremoloRate(v) = 2.5!
    RiffVoiceTremoloDepth(v) = 0.65!
    LStatus "Tremolo/AutoPan 2/4 - Slow tremolo: volume pulses"
    LWait 5

    RiffVoiceTremoloRate(v) = 12!
    RiffVoiceTremoloDepth(v) = 0.9!
    LStatus "Tremolo/AutoPan 3/4 - Fast tremolo: machine-gun/helicopter pulse"
    LWait 5

    RiffVoiceTremoloDepth(v) = 0!
    RiffVoiceAutoPanRate(v) = 0.45!
    RiffVoiceAutoPanDepth(v) = 1!
    LStatus "Tremolo/AutoPan 4/4 - AutoPan: sound moves left/right"
    LWait 8

    LStopClean 0.5!
End Sub

Public Sub Demo_RingModMetallic()
    If Not LStartEngine() Then Exit Sub

    Dim v As Long
    v = RiffPlayOscillator(RiffWaveSine, 330!)
    LSetVoice v, 0.12!, 0!

    LStatus "Ring Mod 1/4 - Clean sine tone"
    LWait 3

    RiffVoiceRingModFreq(v) = 45!
    RiffVoiceRingModMix(v) = 0.35!
    LStatus "Ring Mod 2/4 - Low carrier: bell-like wobble"
    LWait 5

    RiffVoiceRingModFreq(v) = 120!
    RiffVoiceRingModMix(v) = 0.7!
    LStatus "Ring Mod 3/4 - Medium carrier: robotic tone"
    LWait 5

    RiffVoiceRingModFreq(v) = 387!
    RiffVoiceRingModMix(v) = 1!
    LStatus "Ring Mod 4/4 - High carrier: metallic/alien sidebands"
    LWait 6

    LStopClean 0.5!
End Sub

Public Sub Demo_PitchFade()
    If Not LStartEngine() Then Exit Sub

    Dim v As Long, i As Long
    v = RiffPlayOscillator(RiffWaveSawtooth, 220!)
    LSetVoice v, 0!, 0!

    LStatus "Pitch/Fade 1/4 - Fade in"
    For i = 0 To 100 Step 5
        RiffVoiceVolume(v) = (i / 100!) * 0.11!
        LWait 0.08!
    Next i
    LWait 2

    LStatus "Pitch/Fade 2/4 - Pitch bends down"
    For i = 100 To 50 Step -2
        RiffVoicePitch(v) = i / 100!
        LWait 0.07!
    Next i
    LWait 1

    LStatus "Pitch/Fade 3/4 - Pitch bends up above normal"
    For i = 50 To 150 Step 2
        RiffVoicePitch(v) = i / 100!
        LWait 0.055!
    Next i
    LWait 1

    LStatus "Pitch/Fade 4/4 - Fade out"
    For i = 100 To 0 Step -5
        RiffVoiceVolume(v) = (i / 100!) * 0.11!
        LWait 0.08!
    Next i

    LStopClean 0.2!
End Sub

Public Sub Demo_PresetsNoise()
    If Not LStartEngine() Then Exit Sub

    Dim v As Long
    v = RiffPlayOscillator(RiffWaveSine, 220!)
    LSetVoice v, 0.15!, 0!

    LStatus "Presets 1/3 - Sine wave: Radio preset"
    RiffVoiceApplyPreset v, RiffFxRadio
    LWait 4

    LStatus "Presets 2/3 - Sine wave: Underwater preset"
    RiffVoiceApplyPreset v, RiffFxUnderwater
    LWait 4

    LStatus "Presets 3/3 - Sine wave: Robot preset"
    RiffVoiceApplyPreset v, RiffFxRobot
    LWait 4

    RiffStop v
    LWait 0.2!

    LStatus "Noise 1/3 - White Noise: Classic static"
    v = RiffPlayNoise(RiffWaveWhiteNoise)
    LSetVoice v, 0.08!, 0!
    LWait 3
    RiffStop v

    LStatus "Noise 2/3 - Pink Noise: Natural/Rain-like"
    v = RiffPlayNoise(RiffWavePinkNoise)
    LSetVoice v, 0.12!, 0!
    LWait 3
    RiffStop v

    LStatus "Noise 3/3 - Brown Noise: Deep/Thunder-like"
    v = RiffPlayNoise(RiffWaveBrownNoise)
    LSetVoice v, 0.18!, 0!
    LWait 4

    LStopClean 0.5!
End Sub
