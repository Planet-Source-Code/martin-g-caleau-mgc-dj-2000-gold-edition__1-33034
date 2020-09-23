Attribute VB_Name = "Module1"
Public Const HIGHEST_VOLUME_SETTING = 12
Public Const AUX_MAPPER = -1&
Public Const MAXPNAMELEN = 32
Public Movetest As Integer
Type AUXCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
    wTechnology As Integer
    dwSupport As Long
End Type
    Public Const AUXCAPS_CDAUDIO = 1
    Public Const AUXCAPS_AUXIN = 2
    Public Const AUXCAPS_VOLUME = &H1 ' supports volume control
    Public Const AUXCAPS_LRVOLUME = &H2 ' separate left-right volume control
    Type MIXERCONTROLDETAILS
    Fader As Integer
    Volume As Integer
    Bass As Integer
    Treble As Integer
    End Type
Declare Function auxGetNumDevs Lib "winmm.dll" () As Long
Declare Function auxGetDevCaps Lib "winmm.dll" Alias "auxGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As AUXCAPS, ByVal uSize As Long) As Long
Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByRef lpdwVolume As Long) As Long
Declare Function auxOutMessage Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
    Public Const MMSYSERR_NOERROR = 0
    Public Const MMSYSERR_BASE = 0
    Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Type VolumeSetting
    LeftVol As Integer
    RightVol As Integer
End Type

Public Function lSetVolume(ByRef lLeftVol As Long, ByRef lRightVol As Long, lDeviceID As Long) As Long
    Dim bReturnValue As Boolean
    Dim Volume As VolumeSetting
    Dim lAPIReturnVal As Long
    Dim lBothVolumes As Long
    Volume.LeftVol = nSigned(lLeftVol * 65535 / HIGHEST_VOLUME_SETTING)
    Volume.RightVol = nSigned(lRightVol * 65535 / HIGHEST_VOLUME_SETTING)
    lDataLen = Len(Volume)
    CopyMemory lBothVolumes, Volume.LeftVol, lDataLen
    lAPIReturnVal = auxSetVolume(lDeviceID, lBothVolumes)
    lSetVolume = lAPIReturnVal
End Function

Public Function lGetVolume(ByRef lLeftVol As Long, ByRef lRightVol As Long, lDeviceID As Long) As Long
    Dim bReturnValue As Boolean
    Dim Volume As VolumeSetting
    Dim lAPIReturnVal As Long
    Dim lBothVolumes As Long
    lAPIReturnVal = auxGetVolume(lDeviceID, lBothVolumes)
    lDataLen = Len(Volume)
    CopyMemory Volume.LeftVol, lBothVolumes, lDataLen
    lLeftVol = HIGHEST_VOLUME_SETTING * lUnsigned(Volume.LeftVol) / 65535
    lRightVol = HIGHEST_VOLUME_SETTING * lUnsigned(Volume.RightVol) / 65535
    lGetVolume = lAPIReturnVal
End Function

Public Function nSigned(ByVal lUnsignedInt As Long) As Integer
   Dim nReturnVal As Integer
   If lUnsignedInt > 65535 Or lUnsignedInt < 0 Then
        MsgBox "Error in conversion from Unsigned to nSigned Integer"
        nSignedInt = 0
        Exit Function
    End If
        If lUnsignedInt > 32767 Then
            nReturnVal = lUnsignedInt - 65536
        Else
            nReturnVal = lUnsignedInt
       End If
            nSigned = nReturnVal
        End Function

Public Function lUnsigned(ByVal nSignedInt As Integer) As Long
    Dim lReturnVal As Long
    If nSignedInt < 0 Then
        lReturnVal = nSignedInt + 65536
    Else
        lReturnVal = nSignedInt
    End If
        If lReturnVal > 65535 Or lReturnVal < 0 Then
            MsgBox "Error in conversion from nSigned to Unsigned Integer"
            lReturnVal = 0
        End If
            lUnsigned = lReturnVal
        End Function

