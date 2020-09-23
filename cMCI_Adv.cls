VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMCI_Adv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'=============================================================================================================
'
' cMCI_Adv Class Module
' ---------------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone .net )
'
' Last Update : January 8, 2002
'
' VB Versions : 5.0 / 6.0
'
' Requires    : Sound Card
'
' Description : This module allows you to play, pause, stop, and manipulate any multimedia file that you have
'               the drivers installed for.  This module uses existing drivers to play multimedia files.  So
'               if you wanted to use this module to play a QuickTime movie file, you could as long as the
'               user that is running your program has QuickTime installed on their computer.  If you wanted to
'               use this module to play MP3 files, you could as long as the user that is running your program
'               has software installed to play MP3 files (like WinAmp).
'
' NOTE        : I have successfully tested this module on Win9x and WinNT with the following media file formats:
'                 AIF, ASF, AU, AVI, WAV, MID, RMI, AVI, CDA, MP3, MPU, MOV, MPEG, WMA
'
' NOTE        : Specifying "Other" as the device type will always work given the following 3 things are TRUE:
'                 1) The file extention is included as part of the file path (tells it what type of file it is)
'                 2) The file extention is the correct extention for the actual file type
'                 3) There is a multimedia driver installed to handle the specified multimedia type
'
'_____________________________________________________________________________________________________________
' SEE ALSO:
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'
' MCI Overview
' http://msdn.microsoft.com/library/psdk/multimed/mci_7vvt.htm
'
' MCI Reference
' http://msdn.microsoft.com/library/psdk/multimed/mci_25k5.htm
'
' mciSendString Win32 API
' http://msdn.microsoft.com/library/psdk/multimed/mmfunc_5lrb.htm
'
' MCI Error Definitions
' http://msdn.microsoft.com/library/psdk/multimed/mmconst_3stv.htm
'
' MCI Device Types
' http://msdn.microsoft.com/library/psdk/multimed/mci_1k8j.htm
'
' Opening An MCI Device
' http://msdn.microsoft.com/library/psdk/multimed/mci_6oit.htm
'
' mciSendString - "open" Command
' http://msdn.microsoft.com/library/psdk/multimed/mmcmdstr_7226.htm
'
' mciSendString - "notify" Flag
' http://msdn.microsoft.com/library/psdk/multimed/mci_7xyf.htm
'
' MCI Command Strings
' http://msdn.microsoft.com/library/psdk/multimed/mci_04dv.htm
'
' MCI Error - "There is no driver installed on your system"
' http://support.microsoft.com/support/kb/articles/Q178/2/91.ASP
'
' MCI Example - CD Player
' http://support.microsoft.com/support/kb/articles/Q181/4/74.ASP
'
' MCI Example - Play AVI Fullscreen
' http://support.microsoft.com/support/kb/articles/Q171/9/81.asp
'
'=============================================================================================================
'
' LEGAL:
'
' You are free to use this code as long as you keep the above heading information intact and unchanged. Credit
' given where credit is due.  Also, it is not required, but it would be appreciated if you would mention
' somewhere in your compiled program that that your program makes use of code written and distributed by
' Kevin Wilson (www.TheVBZone.com).  Feel free to link to this code via your web site or articles.
'
' You may NOT take this code and pass it off as your own.  You may NOT distribute this code on your own server
' or web site.  You may NOT take code created by Kevin Wilson (www.TheVBZone.com) and use it to create products,
' utilities, or applications that directly compete with products, utilities, and applications created by Kevin
' Wilson, TheVBZone.com, or Wilson Media.  You may NOT take this code and sell it for profit without first
' obtaining the written consent of the author Kevin Wilson.
'
' These conditions are subject to change at the discretion of the owner Kevin Wilson at any time without
' warning or notice.  Copyright© by Kevin Wilson.  All rights reserved.
'
'=============================================================================================================


Private Type MIXERCONTROL
  cbStruct        As Long
  dwControlID     As Long
  dwControlType   As Long
  fdwControl      As Long
  cMultipleItems  As Long
  szShortName     As String * 16 'MIXER_SHORT_NAME_CHARS
  szName          As String * 64 'MIXER_LONG_NAME_CHARS
  lMinimum        As Long
  lMaximum        As Long
  Reserved(10)    As Long
End Type
 
Private Type MIXERLINECONTROLS
  cbStruct        As Long
  dwLineID        As Long
  dwControl       As Long
  cControls       As Long
  cbmxctrl        As Long
  pamxctrl        As Long
End Type

Private Type MIXERLINE
  cbStruct        As Long
  dwDestination   As Long
  dwSource        As Long
  dwLineID        As Long
  fdwLine         As Long
  dwUser          As Long
  dwComponentType As Long
  cChannels       As Long
  cConnections    As Long
  cControls       As Long
  szShortName     As String * 16 'MIXER_SHORT_NAME_CHARS
  szName          As String * 64 'MIXER_LONG_NAME_CHARS
  dwType          As Long
  dwDeviceID      As Long
  wMid            As Integer
  wPid            As Integer
  vDriverVersion  As Long
  szPname         As String * 32 'MAXPNAMELEN
End Type

Private Type MIXERCONTROLDETAILS
  cbStruct       As Long
  dwControlID    As Long
  cChannels      As Long
  item           As Long
  cbDetails      As Long
  paDetails      As Long
End Type

'Private Type MIXERCONTROLDETAILS
'  cbStruct        As Long 'DWORD
'  dwControlID     As Long 'DWORD
'  cChannels       As Long 'DWORD
'  hwndOwner       As Long 'HWND
'  cMultipleItems  As Long 'DWORD
'  cbDetails       As Long 'DWORD
'  paDetails       As Long 'LPVOID
'End Type

Private Type MIXERCONTROLDETAILS_UNSIGNED
  dwValue         As Long 'DWORD
End Type

Public Enum ResDataTypes
  rdt_Bitmap = 0
  rdt_Icon = 1
  rdt_Cursor = 2
  rdt_String = 3
  rdt_Data = 4
End Enum

'_____________________________________________________________________________________________________________
' Device Types Defined In The MSDN:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' CDAudio        CD audio player
' DAT            Digital-audio tape player
' DigitalVideo   Digital video in a window (not GDI-based)
' AVIVideo       (Same as digitalvideo)
' Overlay        Overlay device (analog video in a window)
' Scanner        Image scanner
' Sequencer      MIDI sequencer
' VCR            Video-cassette recorder or player
' VideoDisc      Videodisc player
' WaveAudio      Audio device that plays digitized waveform files
' MPEGVideo      ActiveMovie/DirectShow player
' MPEGVideo2     ActiveMovie/DirectShow player
' Other          Undefined MCI device
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Public Enum MciDeviceTypes ' See Also : http://msdn.microsoft.com/library/psdk/multimed/mci_8vxv.htm
  
  'CD Audio Device          (Basic MCI     - Exists if a sound card has been installed on the user's computer)
  CDAudio = 1               'mcicda.drv    - CDA
  
  'Motion Video Device      (Basic MCI     - Exists if a sound card has been installed on the user's computer)
  '(NOTE : AVI-Video is also known as "Digital-Video")
  AVIVideo = 2              'mciavi.drv    - AVI
  
  'MIDI Sequencer Device    (Basic MCI     - Exists if a sound card has been installed on the user's computer)
  Sequencer = 3             'mciseq.drv    - RMI, MID, MIDI
  
  'Wave Audio Device        (Basic MCI     - Exists if a sound card has been installed on the user's computer)
  WaveAudio = 4             'mciwave.drv   - WAV, WAVE
  
  'ActiveMovie MCI Driver   (Selective MCI - Exists if ActiveMovie/DirectShow have been installed on the user's computer)
  '(NOTE - If MSIE 4.01 or better have been installed on the user's computer, ActiveMovie was installed with it)
  MPEGVideo = 5             'mciqtz.drv / mciqtz32.dll - AIF, AIFC, AIFF, AU, DAT, M1V, MOV, M3U, MOV, MP2, MP2V, MP3, MPA, MPE, MPEG, MPG, MPV2, QT, SND
  MPEGVideo2 = 6            'mciqtz32.dll  - ASF, ASX, IVF, RA, RAM, RM, RMM, WAX, WM, WMA, WMV, WVX
  
  'VISCA VCR Device         (Selective MCI - Exists if a VCR device or VCR software has been installed on the user's computer)
  VCR = 7                   'mcivisca.drv
  
  'PIONEER LaserDisc Device (Selective MCI - Exists if a LaserDisk/VideoDisc device or software has been installed on the user's computer)
  VideoDisc = 8             'mcipionr.drv
  
  'Overlay device - Analog video in a window
  Overlay = 9
  
  ' Other - Type is selected by the file extention
  other = 10
End Enum


'--------------------------------------------
' WAV Capture sizes based on replay settings:
'--------------------------------------------
'  SPS:      BPS:   Channels:  Size:
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'  8.000kHz  8bit   Mono    =  7   k/sec
'  8.000kHz  8bit   Stereo  =  15  k/sec
'  8.000kHz  16bit  Mono    =  15  k/sec
'  8.000kHz  16bit  Stereo  =  31  k/sec
' 11.025kHz  8bit   Mono    =  10  k/sec
' 11.025kHz  8bit   Stereo  =  21  k/sec
' 11.025kHz  16bit  Mono    =  21  k/sec
' 11.025kHz  16bit  Stereo  =  43  k/sec  <-- GOOD DEFAULT
'  22.050Hz  8bit   Mono    =  21  k/sec
'  22.050Hz  8bit   Stereo  =  43  k/sec
'  22.050Hz  16bit  Mono    =  43  k/sec
'  22.050Hz  16bit  Stereo  =  86  k/sec
'  44.100Hz  8bit   Mono    =  43  k/sec
'  44.100Hz  8bit   Stereo  =  86  k/sec
'  44.100Hz  16bit  Mono    =  86  k/sec
'  44.100Hz  16bit  Stereo  =  172 k/sec
'--------------------------------------------

Public Enum BitsPerSample
  bps_8bit = 8
  bps_16bit = 16
End Enum

Public Enum Channels
  Mono = 1
  Stereo = 2
End Enum

Public Enum SamplesPerSec
  sps_8000 = 8000   ' 8.000kHz
  sps_11025 = 11025 ' 11.025kHz
  sps_22050 = 22050 ' 22.050Hz
  sps_44100 = 44100 ' 44.100Hz
End Enum

' Enumeration - MIXERLINE.dwComponentType
Private Enum MixerComponentTypes
  MIXERLINE_COMPONENTTYPE_DST_DIGITAL = 1        ' Audio line is a digital destination (for example, digital input to a DAT or CD audio device).
  MIXERLINE_COMPONENTTYPE_DST_HEADPHONES = 5     ' Audio line is an adjustable (gain and/or attenuation) destination intended to drive headphones. Most audio cards use the same audio destination line for speakers and headphones, in which case the mixer device simply uses the MIXERLINE_COMPONENTTYPE_DST_SPEAKERS type.
  MIXERLINE_COMPONENTTYPE_DST_LINE = 2           ' Audio line is a line level destination (for example, line level input from a CD audio device) that will be the final recording source for the analog-to-digital converter (ADC). Because most audio cards for personal computers provide some sort of gain for the recording audio source line, the mixer device will use the MIXERLINE_COMPONENTTYPE_DST_WAVEIN type.
  MIXERLINE_COMPONENTTYPE_DST_MONITOR = 3        ' Audio line is a destination used for a monitor.
  MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = 4       ' Audio line is an adjustable (gain and/or attenuation) destination intended to drive speakers. This is the typical component type for the audio output of audio cards for personal computers.
  MIXERLINE_COMPONENTTYPE_DST_TELEPHONE = 6      ' Audio line is a destination that will be routed to a telephone line.
  MIXERLINE_COMPONENTTYPE_DST_UNDEFINED = 0      ' Audio line is a destination that cannot be defined by one of the standard component types. A mixer device is required to use this component type for line component types that have not been defined by Microsoft Corporation.
  MIXERLINE_COMPONENTTYPE_DST_VOICEIN = 8        ' Audio line is a destination that will be the final recording source for voice input. This component type is exactly like MIXERLINE_COMPONENTTYPE_DST_WAVEIN but is intended specifically for settings used during voice recording/recognition. Support for this line is optional for a mixer device. Many mixer devices provide only MIXERLINE_COMPONENTTYPE_DST_WAVEIN.
  MIXERLINE_COMPONENTTYPE_DST_WAVEIN = 7         ' Audio line is a destination that will be the final recording source for the waveform-audio input (ADC). This line typically provides some sort of gain or attenuation. This is the typical component type for the recording line of most audio cards for personal computers.
  MIXERLINE_COMPONENTTYPE_SRC_ANALOG = 4106      ' Audio line is an analog source (for example, analog output from a video-cassette tape).
  MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY = 4105   ' Audio line is a source originating from the auxiliary audio line. This line type is intended as a source with gain or attenuation that can be routed to the MIXERLINE_COMPONENTTYPE_DST_SPEAKERS destination and/or recorded from the MIXERLINE_COMPONENTTYPE_DST_WAVEIN destination.
  MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = 4101 ' Audio line is a source originating from the output of an internal audio CD. This component type is provided for audio cards that provide an audio source line intended to be connected to an audio CD (or CD-ROM playing an audio CD).
  MIXERLINE_COMPONENTTYPE_SRC_DIGITAL = 4097     ' Audio line is a digital source (for example, digital output from a DAT or audio CD).
  MIXERLINE_COMPONENTTYPE_SRC_LINE = 4098        ' Audio line is a line-level source (for example, line-level input from an external stereo) that can be used as an optional recording source. Because most audio cards for personal computers provide some sort of gain for the recording source line, the mixer device will use the MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY type.
  MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = 4099  ' Audio line is a microphone recording source. Most audio cards for personal computers provide at least two types of recording sources: an auxiliary audio line and microphone input. A microphone audio line typically provides some sort of gain. Audio cards that use a single input for use with a microphone or auxiliary audio line should use the MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE component type.
  MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER = 4103   ' Audio line is a source originating from personal computer speaker. Several audio cards for personal computers provide the ability to mix what would typically be played on the internal speaker with the output of an audio card. Some audio cards support the ability to use this output as a recording source.
  MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = 4100 ' Audio line is a source originating from the output of an internal synthesizer. Most audio cards for personal computers provide some sort of MIDI synthesizer (for example, an Adlib®-compatible or OPL/3 FM synthesizer).
  MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE = 4102   ' Audio line is a source originating from an incoming telephone line.
  MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED = 4096   ' Audio line is a source that cannot be defined by one of the standard component types. A mixer device is required to use this component type for line component types that have not been defined by Microsoft Corporation.
  MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = 4104     ' Audio line is a source originating from the waveform-audio output digital-to-analog converter (DAC). Most audio cards for personal computers provide this component type as a source to the MIXERLINE_COMPONENTTYPE_DST_SPEAKERS destination. Some cards also allow this source to be routed to the MIXERLINE_COMPONENTTYPE_DST_WAVEIN destination.
End Enum

' Constants - Window Sytles (Used in the mciSendString "open" command for the "Style_Type" parameter)
Private Const WS_BORDER = &H800000                 ' Creates a window that has a thin-line border.
Private Const WS_CAPTION = &HC00000                ' Creates a window that has a title bar (includes the WS_BORDER style).
Private Const WS_CHILD = &H40000000                ' Creates a child window. A window with this style cannot have a menu bar. This style cannot be used with the WS_POPUP style.
Private Const WS_CLIPCHILDREN = &H2000000          ' Excludes the area occupied by child windows when drawing occurs within the parent window. This style is used when creating the parent window.
Private Const WS_CLIPSIBLINGS = &H4000000          ' Clips child windows relative to each other; that is, when a particular child window receives a WM_PAINT message, the WS_CLIPSIBLINGS style clips all other overlapping child windows out of the region of the child window to be updated. If WS_CLIPSIBLINGS is not specified and child windows overlap, it is possible, when drawing within the client area of a child window, to draw within the client area of a neighboring child window.
Private Const WS_DISABLED = &H8000000              ' Creates a window that is initially disabled. A disabled window cannot receive input from the user. To change this after a window has been created, use EnableWindow.
Private Const WS_DLGFRAME = &H400000               ' Creates a window that has a border of a style typically used with dialog boxes. A window with this style cannot have a title bar.
Private Const WS_GROUP = &H20000                   ' Specifies the first control of a group of controls. The group consists of this first control and all controls defined after it, up to the next control with the WS_GROUP style. The first control in each group usually has the WS_TABSTOP style so that the user can move from group to group. The user can subsequently change the keyboard focus from one control in the group to the next control in the group by using the direction keys.  You can turn this style on and off to change dialog box navigation. To change this style after a window has been created, use SetWindowLong.
Private Const WS_HSCROLL = &H100000                ' Creates a window that has a horizontal scroll bar.
Private Const WS_MAXIMIZE = &H1000000              ' Creates a window that is initially maximized.
Private Const WS_MAXIMIZEBOX = &H10000             ' Creates a window that has a maximize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.
Private Const WS_MINIMIZE = &H20000000             ' Creates a window that is initially minimized. Same as the WS_ICONIC style.
Private Const WS_MINIMIZEBOX = &H20000             ' Creates a window that has a minimize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.
Private Const WS_ICONIC = WS_MINIMIZE              ' Creates a window that is initially minimized. Same as the WS_MINIMIZE style.
Private Const WS_OVERLAPPED = &H0                  ' Creates an overlapped window. An overlapped window has a title bar and a border. Same as the WS_TILED style.
Private Const WS_POPUP = &H80000000                ' Creates a pop-up window. This style cannot be used with the WS_CHILD style.
Private Const WS_SIZEBOX = &H40000                 ' Creates a window that has a sizing border. Same as the WS_THICKFRAME style.
Private Const WS_SYSMENU = &H80000                 ' Creates a window that has a window menu on its title bar. The WS_CAPTION style must also be specified.
Private Const WS_TABSTOP = &H10000                 ' Specifies a control that can receive the keyboard focus when the user presses the TAB key. Pressing the TAB key changes the keyboard focus to the next control with the WS_TABSTOP style.  You can turn this style on and off to change dialog box navigation. To change this style after a window has been created, use SetWindowLong.
Private Const WS_THICKFRAME = &H40000              ' Creates a window that has a sizing border. Same as the WS_SIZEBOX style.
Private Const WS_VISIBLE = &H10000000              ' Creates a window that is initially visible.  This style can be turned on and off by using ShowWindow or SetWindowPos.
Private Const WS_VSCROLL = &H200000                ' Creates a window that has a vertical scroll bar.
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX) ' Creates an overlapped window with the WS_OVERLAPPED, WS_CAPTION, WS_SYSMENU, WS_THICKFRAME, WS_MINIMIZEBOX, and WS_MAXIMIZEBOX styles. Same as the WS_TILEDWINDOW style.
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU) ' Creates a pop-up window with WS_BORDER, WS_POPUP, and WS_SYSMENU styles. The WS_CAPTION and WS_POPUPWINDOW styles must be combined to make the window menu visible.
Private Const WS_CHILDWINDOW = WS_CHILD            ' Same as the WS_CHILD style.
Private Const WS_TILED = WS_OVERLAPPED             ' Creates an overlapped window. An overlapped window has a title bar and a border. Same as the WS_OVERLAPPED style.
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW ' Creates an overlapped window with the WS_OVERLAPPED, WS_CAPTION, WS_SYSMENU, WS_THICKFRAME, WS_MINIMIZEBOX, and WS_MAXIMIZEBOX styles. Same as the WS_OVERLAPPEDWINDOW style.

' Error Constants
Private Const MMSYSERR_NOERROR = 0
Private Const MCIERR_BASE = 256
Private Const MCIERR_BAD_CONSTANT = (MCIERR_BASE + 34)
Private Const MCIERR_BAD_INTEGER = (MCIERR_BASE + 14)
Private Const MCIERR_BAD_TIME_FORMAT = (MCIERR_BASE + 37)
Private Const MCIERR_CANNOT_LOAD_DRIVER = (MCIERR_BASE + 10)
Private Const MCIERR_CANNOT_USE_ALL = (MCIERR_BASE + 23)
Private Const MCIERR_CREATEWINDOW = (MCIERR_BASE + 91)
Private Const MCIERR_CUSTOM_DRIVER_BASE = (MCIERR_BASE + 256)
Private Const MCIERR_DEVICE_LENGTH = (MCIERR_BASE + 54)
Private Const MCIERR_DEVICE_LOCKED = (MCIERR_BASE + 32)
Private Const MCIERR_DEVICE_NOT_INSTALLED = (MCIERR_BASE + 50)
Private Const MCIERR_DEVICE_NOT_READY = (MCIERR_BASE + 20)
Private Const MCIERR_DEVICE_OPEN = (MCIERR_BASE + 9)
Private Const MCIERR_DEVICE_ORD_LENGTH = (MCIERR_BASE + 55)
Private Const MCIERR_DEVICE_TYPE_REQUIRED = (MCIERR_BASE + 31)
Private Const MCIERR_DRIVER = (MCIERR_BASE + 22)
Private Const MCIERR_DRIVER_INTERNAL = (MCIERR_BASE + 16)
Private Const MCIERR_DUPLICATE_ALIAS = (MCIERR_BASE + 33)
Private Const MCIERR_DUPLICATE_FLAGS = (MCIERR_BASE + 39)
Private Const MCIERR_EXTENSION_NOT_FOUND = (MCIERR_BASE + 25)
Private Const MCIERR_EXTRA_CHARACTERS = (MCIERR_BASE + 49)
Private Const MCIERR_FILE_NOT_FOUND = (MCIERR_BASE + 19)
Private Const MCIERR_FILE_NOT_SAVED = (MCIERR_BASE + 30)
Private Const MCIERR_FILE_READ = (MCIERR_BASE + 92)
Private Const MCIERR_FILE_WRITE = (MCIERR_BASE + 93)
Private Const MCIERR_FILENAME_REQUIRED = (MCIERR_BASE + 48)
Private Const MCIERR_FLAGS_NOT_COMPATIBLE = (MCIERR_BASE + 28)
Private Const MCIERR_GET_CD = (MCIERR_BASE + 51)
Private Const MCIERR_HARDWARE = (MCIERR_BASE + 6)
Private Const MCIERR_ILLEGAL_FOR_AUTO_OPEN = (MCIERR_BASE + 47)
Private Const MCIERR_INTERNAL = (MCIERR_BASE + 21)
Private Const MCIERR_INVALID_DEVICE_ID = (MCIERR_BASE + 1)
Private Const MCIERR_INVALID_DEVICE_NAME = (MCIERR_BASE + 7)
Private Const MCIERR_INVALID_FILE = (MCIERR_BASE + 40)
Private Const MCIERR_MISSING_COMMAND_STRING = (MCIERR_BASE + 11)
Private Const MCIERR_MISSING_DEVICE_NAME = (MCIERR_BASE + 36)
Private Const MCIERR_MISSING_PARAMETER = (MCIERR_BASE + 17)
Private Const MCIERR_MISSING_STRING_ARGUMENT = (MCIERR_BASE + 13)
Private Const MCIERR_MULTIPLE = (MCIERR_BASE + 24)
Private Const MCIERR_MUST_USE_SHAREABLE = (MCIERR_BASE + 35)
Private Const MCIERR_NEW_REQUIRES_ALIAS = (MCIERR_BASE + 43)
Private Const MCIERR_NO_CLOSING_QUOTE = (MCIERR_BASE + 38)
Private Const MCIERR_NO_ELEMENT_ALLOWED = (MCIERR_BASE + 45)
Private Const MCIERR_NO_INTEGER = (MCIERR_BASE + 56)
Private Const MCIERR_NO_WINDOW = (MCIERR_BASE + 90)
Private Const MCIERR_NONAPPLICABLE_FUNCTION = (MCIERR_BASE + 46)
Private Const MCIERR_NOTIFY_ON_AUTO_OPEN = (MCIERR_BASE + 44)
Private Const MCIERR_NULL_PARAMETER_BLOCK = (MCIERR_BASE + 41)
Private Const MCIERR_OUT_OF_MEMORY = (MCIERR_BASE + 8)
Private Const MCIERR_OUTOFRANGE = (MCIERR_BASE + 26)
Private Const MCIERR_PARAM_OVERFLOW = (MCIERR_BASE + 12)
Private Const MCIERR_PARSER_INTERNAL = (MCIERR_BASE + 15)
Private Const MCIERR_SEQ_DIV_INCOMPATIBLE = (MCIERR_BASE + 80)
Private Const MCIERR_SEQ_NOMIDIPRESENT = (MCIERR_BASE + 87)
Private Const MCIERR_SEQ_PORT_INUSE = (MCIERR_BASE + 81)
Private Const MCIERR_SEQ_PORT_MAPNODEVICE = (MCIERR_BASE + 83)
Private Const MCIERR_SEQ_PORT_MISCERROR = (MCIERR_BASE + 84)
Private Const MCIERR_SEQ_PORT_NONEXISTENT = (MCIERR_BASE + 82)
Private Const MCIERR_SEQ_PORTUNSPECIFIED = (MCIERR_BASE + 86)
Private Const MCIERR_SEQ_TIMER = (MCIERR_BASE + 85)
Private Const MCIERR_SET_CD = (MCIERR_BASE + 52)
Private Const MCIERR_SET_DRIVE = (MCIERR_BASE + 53)
Private Const MCIERR_UNNAMED_RESOURCE = (MCIERR_BASE + 42)
Private Const MCIERR_UNRECOGNIZED_COMMAND = (MCIERR_BASE + 5)
Private Const MCIERR_UNRECOGNIZED_KEYWORD = (MCIERR_BASE + 3)
Private Const MCIERR_UNSUPPORTED_FUNCTION = (MCIERR_BASE + 18)
Private Const MCIERR_WAVE_INPUTSINUSE = (MCIERR_BASE + 66)
Private Const MCIERR_WAVE_INPUTSUNSUITABLE = (MCIERR_BASE + 72)
Private Const MCIERR_WAVE_INPUTUNSPECIFIED = (MCIERR_BASE + 69)
Private Const MCIERR_WAVE_OUTPUTSINUSE = (MCIERR_BASE + 64)
Private Const MCIERR_WAVE_OUTPUTSUNSUITABLE = (MCIERR_BASE + 70)
Private Const MCIERR_WAVE_OUTPUTUNSPECIFIED = (MCIERR_BASE + 68)
Private Const MCIERR_WAVE_SETINPUTINUSE = (MCIERR_BASE + 67)
Private Const MCIERR_WAVE_SETINPUTUNSUITABLE = (MCIERR_BASE + 73)
Private Const MCIERR_WAVE_SETOUTPUTINUSE = (MCIERR_BASE + 65)
Private Const MCIERR_WAVE_SETOUTPUTUNSUITABLE = (MCIERR_BASE + 71)

' Volume Constants
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2
Private Const MIXER_GETCONTROLDETAILSF_VALUE = &H0
Private Const MIXER_SETCONTROLDETAILSF_VALUE = &H0
Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Private Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)

' Drive Identification Constants
Private Const DRIVE_CDROM = 5
Private Const DRIVE_FIXED = 3
Private Const DRIVE_RAMDISK = 6
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_UNKNOWN = 0

' Constants - SendMessage
Private Const WM_SETICON = &H80
Private Const ICON_BIG = 1

' Private Variable Declarations
Private bUserCanceled As Boolean

' Volume Related Variables
Private VolCtrl_Sys    As MIXERCONTROL
Private VolCtrl_Wav    As MIXERCONTROL
Private VolCtrl_Mid    As MIXERCONTROL
Private VolCtrl_CD     As MIXERCONTROL
Private VolMin_Sys     As Long
Private VolMax_Sys     As Long
Private VolMin_Wav     As Long
Private VolMax_Wav     As Long
Private VolMin_Mid     As Long
Private VolMax_Mid     As Long
Private VolMin_CD      As Long
Private VolMax_CD      As Long
Private hMixer         As Long

' Win32 API Function Declarations
Private Declare Sub CopyPtrFromStruct Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Private Declare Sub CopyStructFromPtr Lib "kernel32.dll" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Private Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByRef lpdwVolume As Long) As Long
Private Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function GetDriveTypeA Lib "kernel32.dll" (ByVal nDrive As String) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal strLongPath As String, ByVal strReturnBuffer As String, ByVal ReturnBufferLen As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal ErrorNumber As Long, ByVal ReturnBuffer As String, ByVal ReturnBufferSize As Long) As Long 'BOOL
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal CommandString As String, ByVal ReturnBuffer As String, ByVal ReturnBufferSize As Long, ByVal hCallback As Long) As Long 'MCIERROR
Private Declare Function mixerClose Lib "winmm.dll" (ByVal hMixer As Long) As Long
Private Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hMixerObj As Long, ByRef pMixerCD As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hMixerObj As Long, pMixerLc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Private Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hMixerObj As Long, pMixerL As MIXERLINE, ByVal fdwInfo As Long) As Long
Private Declare Function mixerOpen Lib "winmm.dll" (ByRef phMixer As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hMixerObj As Long, ByRef pMixerCD As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function SetWindowText Lib "USER32" (ByVal hWnd As Long, ByVal LPString As String)
Private Declare Function SendMessage Lib "USER32.DLL" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



Private Sub Class_Initialize()
On Error Resume Next
  
  MCI_CloseAll
  
  ' Open the volume mixer with deviceID 0
  If mixerOpen(hMixer, 0, 0, 0, 0) <> MMSYSERR_NOERROR Then
    MsgBox "An error occured while trying to access the volume mixer.", vbOKOnly + vbExclamation, "  Error Opening Volume Mixer"
  Else
    
    ' Get SYSTEM VOLUME information
    If GetVolumeObject(hMixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, VolCtrl_Sys) = True Then
      VolMax_Sys = VolCtrl_Sys.lMaximum
      VolMin_Sys = VolCtrl_Sys.lMinimum
    End If
    
    ' Get WAVE VOLUME information
    If GetVolumeObject(hMixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, VolCtrl_Wav) = True Then
      VolMax_Wav = VolCtrl_Wav.lMaximum
      VolMin_Wav = VolCtrl_Wav.lMinimum
    End If
    
    ' Get MIDI VOLUME information
    If GetVolumeObject(hMixer, MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER, VolCtrl_Mid) = True Then
      VolMax_Mid = VolCtrl_Mid.lMaximum
      VolMin_Mid = VolCtrl_Mid.lMinimum
    End If
    
    ' Get CD-ROM VOLUME information
    If GetVolumeObject(hMixer, MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC, VolCtrl_CD) = True Then
      VolMax_CD = VolCtrl_CD.lMaximum
      VolMin_CD = VolCtrl_CD.lMinimum
    End If
  End If
  
End Sub

Private Sub Class_Terminate()
On Error Resume Next
  
  MCI_CloseAll
  mixerClose hMixer
  
End Sub



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



'=============================================================================================================
' MCI_Close
'
' Purpose :
' Closes the specified device, and frees up any memory associated with that device.
'
' Param                Use
' ------------------------------------
' DeviceAlias          Specifies the name of the device to perform the action on.
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' WaitToComplete       Optional. If this is set to TRUE, this function will not return control to the application
'                      until the MCI call has completed.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCI_Close(ByVal DeviceAlias As String, Optional ByVal hCallbackProc As Long, Optional ByVal WaitToComplete As Boolean = True, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  ' Make sure parameters are valid
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  
  ' Close the device
  MCI_Close = MCIa_SendCommand("close " & DeviceAlias, , , hCallbackProc, WaitToComplete, ShowErrors)
  
End Function

'=============================================================================================================
' MCI_CloseAll
'
' Purpose :
' Closes all devices opened by the current program, and frees up any memory associated with those device.
'
' Param                Use
' ------------------------------------
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' WaitToComplete       Optional. If this is set to TRUE, this function will not return control to the application
'                      until the MCI call has completed.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCI_CloseAll(Optional ByVal hCallbackProc As Long, Optional ByVal WaitToComplete As Boolean = True, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  ' Close all devices opened by this program
  MCI_CloseAll = MCIa_SendCommand("close all", , , hCallbackProc, WaitToComplete, ShowErrors)
  
End Function

'=============================================================================================================
' MCI_Door
'
' Purpose :
' Either opens or closes the door to the specified device.  This is most commonly used for CD-ROM drives, but
' can be used for other device types as well.
'
' Param                Use
' ------------------------------------
' DeviceAlias          Specifies the name of the device to perform the action on.
' OpenDoor             If set to true, opens the door.  If set to FALSE, closes the door.
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' WaitToComplete       Optional. If this is set to TRUE, this function will not return control to the application
'                      until the MCI call has completed.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCI_Door(ByVal DeviceAlias As String, ByVal OpenDoor As Boolean, Optional ByVal hCallbackProc As Long, Optional ByVal WaitToComplete As Boolean = True, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  ' Make sure parameters are valid
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  
  ' Open or close the door to the device
  If OpenDoor = True Then
    MCI_Door = MCIa_SendCommand("Set " & DeviceAlias & " door open", , , hCallbackProc, WaitToComplete, ShowErrors)
  Else
    MCI_Door = MCIa_SendCommand("Set " & DeviceAlias & " door closed", , , hCallbackProc, WaitToComplete, ShowErrors)
  End If
  
End Function

'=============================================================================================================
' MCI_Open
'
' Purpose :
' Opens the specified device and assigns it the name, or alias, that you give it.  After calling OPEN, you
' can then make other calls to it.
'
' *NOTE - If you specify "MPEGVideo" as the DeviceType parameter, it will most likely play all common media
'         types.  In fact, it plays AVI files better than the "AVIVideo" device type most of the time.
'         However, this option is not available if ActiveMove/DirectShow are not installed.
'
' *NOTE - The "Shareable" option is not very widely supported.  If you get an error about an option not being
'         supported, try taking off that option.
'
' *NOTE - This function makes the application wait for the device to open because it must be open to make
'         further calls to it successfully
'
' Param                Use
' ------------------------------------
' DeviceAlias          Specifies the name of the device to perform the action on.
' DeviceType           Optional. Specifies the type of device to open.  The default is "other".  If left
'                      at "other", chances are the device will work just fine because "other" opens just about
'                      any device type.
' FilePath             Optional. Specifies the path to the file to open.  This is optional because it is
'                      possible to open devices without specifying a source for the device... such as when you
'                      open a new device to write out to.
' Shareable            Optional. If this is set to TRUE, other devices aliases will be able to access the
'                      same device at the same time.  This option is NOT very widely supported among multimedia
'                      device drivers so it is recommended that you do not use it.
' BufferSize           Optional. Sets the size, in seconds, of the buffer used by the waveform-audio device.
'                      The default size of the buffer is set when the waveform-audio device is installed or
'                      configured. Typically the buffer size is set to 4 seconds. With the MCIWAVE device, the
'                      minimum size is 2 seconds and the maximum size is 9 seconds.
' Parent_Handle        Optional. Specifies the window handle of the parent window.  This applies to AVIVideo,
'                      MPEGVideo, MPEGVideo2, and Overlay device types only.
' Style_Child          Optional. If set to true, the AVIVideo, MPEGVideo, MPEGVideo2, or Overlay video being
'                      played will play within the parent specified in the "Parent_Handle" parameter.  So if
'                      you set the "Parent_Handle" to the handle of a PictureBox control and set this parameter
'                      to TRUE, the video will be displayed within that PictureBox control instead of in a
'                      new popup window.
' Style_Overlapped     Optional. If set to TRUE, opens a window with an overlapped window style.  This only
'                      applies to AVIVideo, MPEGVideo, MPEGVideo2, or Overlay device types.
' Style_Popup          Optional. If set to TRUE, opens a window with a pop-up window style.  This only
'                      applies to AVIVideo, MPEGVideo, MPEGVideo2, or Overlay device types.
' Style_Type           Optional. Is not really necissary but allows you to specify additional window styles
'                      via the WS_* constant values (i.e. - WS_BORDER, etc)
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'-------------------------------------------------------------------------------------------------------------
'
' The following table lists the flags that can be specified in the lpszOpenFlags parameter and their meanings:
'
' Open statement options:    Meaning:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' type <device_type>         Specifies the device type of a file.
' alias <device_alias>       Specifies an alternate name for the given device.
'                            If specified, it must be used as the device_id in subsequent commands.
' shareable                  Initializes the device or file as shareable. Subsequent attempts to open the
'                            device or file fail unless you specify "shareable" in both the original and
'                            subsequent open commands.  MCI returns an invalid device error if the device
'                            is already open and not shareable.  The MCISEQ sequencer and MCIWAVE devices
'                            do not support shared files.
' <elementname>              Specifies the name of the device element (file) loaded when the device opens.
' buffer <buffer_size>       Sets the size, in seconds, of the buffer used by the waveform-audio device.
'                            The default size of the buffer is set when the waveform-audio device is
'                            installed or configured. Typically the buffer size is set to 4 seconds.
'                            With the MCIWAVE device, the minimum size is 2 seconds and the maximum size
'                            is 9 seconds.
' parent <hWnd>              Specifies the window handle of the parent window.
' style child                Opens a window with a child window style.
' style overlapped           Opens a window with an overlapped window style.
' style popup                Opens a window with a pop-up window style.
' style <style_type>         Indicates a window style.
'
'=============================================================================================================
Public Function MCI_Open(ByVal DeviceAlias As String, Optional ByVal DeviceType As MciDeviceTypes = other, Optional ByVal FilePath As String, Optional ByVal Shareable As Boolean = False, Optional ByVal BufferSize As Integer, Optional ByVal Parent_Handle As Long, Optional ByVal Style_Child As Boolean, Optional ByVal Style_Overlapped As Boolean, Optional ByVal Style_Popup As Boolean, Optional ByVal Style_Type As String, Optional ByVal hCallbackProc As Long, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  Dim OpenString As String
  Dim DeviceName As String
  
  ' Make sure parameters are valid
  Style_Type = Trim(Style_Type)
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  FilePath = Trim(FilePath)
  FilePath = Get16bitFilePath(FilePath) 'Path must be in 16bit form (no spaces allowed)
  
  ' Specify the name of the device
  Select Case DeviceType
    Case CDAudio:    DeviceName = "cdaudio"
    Case AVIVideo:   DeviceName = "avivideo"
    Case Sequencer:  DeviceName = "Sequencer"
    Case WaveAudio:  DeviceName = "WaveAudio"
    Case MPEGVideo:  DeviceName = "mpegvideo"
    Case MPEGVideo2: DeviceName = "mpegvideo2"
    Case VCR:        DeviceName = "vcr"
    Case VideoDisc:  DeviceName = "videodisc"
    Case Overlay:    DeviceName = "overlay"
    Case other:      DeviceName = "other"
    Case Else:       DeviceName = "other"
  End Select
  
  ' Create the open string based on the options specified
  Select Case DeviceType
    
    ' Options specific to CD Audio
    Case CDAudio
      If FilePath = "" Then
        OpenString = "open " & DeviceName & " alias " & DeviceAlias & " wait"
      Else
        FilePath = LCase(Left(FilePath, 1)) ' Get the CD-ROM drive letter to use
        If Asc(FilePath) >= 97 And Asc(FilePath) <= 122 Then
          OpenString = "open " & FilePath & ": type " & DeviceName & " alias " & DeviceAlias & " wait"
        Else
          OpenString = "open " & DeviceName & " alias " & DeviceAlias & " wait"
        End If
      End If
      If Shareable = True Then OpenString = OpenString & " shareable"
      
    ' Options that pertain to WaveAudio, Sequencer, VCR, and VideoDisc (non-video / basic options)
    Case WaveAudio, Sequencer, VCR, VideoDisc
      If FilePath <> "" And Dir(FilePath) <> "" Then
        OpenString = "open " & DeviceName & "!" & FilePath & " alias " & DeviceAlias & " wait"
      Else
        OpenString = "open " & DeviceName & " alias " & DeviceAlias & " wait"
      End If
      If DeviceType = WaveAudio Then
        If BufferSize > 0 Then OpenString = OpenString & " buffer " & CStr(BufferSize)
      End If
      If Shareable = True Then OpenString = OpenString & " shareable"
      
    ' Options that pertain to AVIVideo, MPEGVideo, MPEGVideo2, and Overlay (video options)
    Case AVIVideo, MPEGVideo, MPEGVideo2, Overlay
      If FilePath <> "" And Dir(FilePath) <> "" Then
        OpenString = "open " & DeviceName & "!" & FilePath & " alias " & DeviceAlias & " wait"
      Else
        OpenString = "open " & DeviceName & " alias " & DeviceAlias & " wait"
      End If
      If Parent_Handle <> 0 Then OpenString = OpenString & " parent " & CStr(Parent_Handle)
      If Shareable = True Then OpenString = OpenString & " shareable"
      If Style_Child = True Then OpenString = OpenString & " style child"
      If Style_Overlapped = True Then OpenString = OpenString & " style overlapped"
      If Style_Popup = True Then OpenString = OpenString & " style popup"
      If Style_Type <> "" Then OpenString = OpenString & " style " & Style_Type
      
    ' Device = NOT SPECIFIED (selected by the file extention) * NOTE - Not all parameters apply to each file type... see documentation for the "MciDeviceTypes" enumeration for more info
    Case Else
      If FilePath <> "" And Dir(FilePath) <> "" Then
        OpenString = "open " & FilePath & " alias " & DeviceAlias & " wait"
      Else
        Exit Function
      End If
      If Parent_Handle <> 0 Then OpenString = OpenString & " parent " & CStr(Parent_Handle)
      If Shareable = True Then OpenString = OpenString & " shareable"
      If Style_Child = True Then OpenString = OpenString & " style child"
      If Style_Overlapped = True Then OpenString = OpenString & " style overlapped"
      If Style_Popup = True Then OpenString = OpenString & " style popup"
      If Style_Type <> "" Then OpenString = OpenString & " style " & Style_Type
      
  End Select
  
  ' Open the device according to the parameters set
  MCI_Open = MCIa_SendCommand(OpenString, , , hCallbackProc, False, ShowErrors)
  

End Function

'=============================================================================================================
' MCI_Pause
'
' Purpose :
' Pauses the playback of the specified Device.
'
' Param                Use
' ------------------------------------
' DeviceAlias          Specifies the name of the device to perform the action on.
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' WaitToComplete       Optional. If this is set to TRUE, this function will not return control to the application
'                      until the MCI call has completed.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCI_Pause(ByVal DeviceAlias As String, Optional ByVal hCallbackProc As Long, Optional ByVal WaitToComplete As Boolean = False, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  ' Make sure parameters are valid
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  
  ' Pause the device
  MCI_Pause = MCIa_SendCommand("pause " & DeviceAlias, , , hCallbackProc, WaitToComplete, ShowErrors)
  
End Function

'=============================================================================================================
' MCI_Play
'
' Purpose :
' Starts the playback of the specified Device.
'
' *NOTE - The "PlayFullScreen" parameter is only supported by AVIVideo and MPEGVideo
'
' Param                Use
' ------------------------------------
' DeviceAlias          Specifies the name of the device to perform the action on.
' SpecifyStartPosition Optional. If set to TRUE, allows you to specify at what position to start the playback.
' StartPosition        Optional. If the "SpecifyStartPosition" parameter is set to TRUE, this parameter
'                      specifies at what position to start playback at.
' SpecifyEndPosition   Optional. If set to TRUE, allows you to specify at what position to STOP the playback.
' EndPosition          Optional. If the "SpecifyEndPosition" parameter is set to TRUE, this parameter
'                      specifies at what position to stop playback at.
' PlayFullScreen       Optional. When playing video, you can set this parameter to play the video full screen.
'                      Use this only when playing compressed files, uncompressed files won't play full-screen.
' RepeatPlay           Optional. If this parameter is set to TRUE, the device will continue to play in a loop
'                      over and over again instead of stoping at th end of the media.
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' WaitToComplete       Optional. If this is set to TRUE, this function will not return control to the application
'                      until the MCI call has completed.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCI_Play(ByVal DeviceAlias As String, Optional ByVal SpecifyStartPosition As Boolean, Optional ByVal StartPosition As Long, Optional ByVal SpecifyEndPosition As Boolean, Optional ByVal EndPosition As Long, Optional ByVal PlayFullScreen As Boolean, Optional ByVal RepeatPlay As Boolean, Optional ByVal hCallbackProc As Long, Optional ByVal WaitToComplete As Boolean = False, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  Dim PlayString As String
  
  ' Make sure parameters are valid
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  
  ' Assemble the string to be used to play the device
  PlayString = "play " & DeviceAlias
  If SpecifyStartPosition = True Then PlayString = PlayString & " from " & CStr(StartPosition)
  If SpecifyEndPosition = True Then PlayString = PlayString & " to " & CStr(StartPosition)
  If RepeatPlay = True Then PlayString = PlayString & " repeat"
  If PlayFullScreen = True Then PlayString = PlayString & " fullscreen" 'This only works for AVI files
  
  ' Play the device
  MCI_Play = MCIa_SendCommand(PlayString, , , hCallbackProc, WaitToComplete, ShowErrors)
  
End Function

'=============================================================================================================
' MCI_Resume
'
' Purpose :
' Resumes the play of paused devices.
'
' Param                Use
' ------------------------------------
' DeviceAlias          Specifies the name of the device to perform the action on.
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' WaitToComplete       Optional. If this is set to TRUE, this function will not return control to the application
'                      until the MCI call has completed.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCI_Resume(ByVal DeviceAlias As String, Optional ByVal hCallbackProc As Long, Optional ByVal WaitToComplete As Boolean = False, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  ' Make sure parameters are valid
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  
  ' Resume the device
  MCI_Resume = MCIa_SendCommand("resume " & DeviceAlias, , , hCallbackProc, WaitToComplete, ShowErrors)
  
End Function

'=============================================================================================================
' MCI_Seek
'
' Purpose :
' Skips the current position of the playback ahead (or back) to the specified seek position.
'
' Param                Use
' ------------------------------------
' DeviceAlias          Specifies the name of the device to perform the action on.
' SeekPosistion        Specifies the seek position to skip to.
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' WaitToComplete       Optional. If this is set to TRUE, this function will not return control to the application
'                      until the MCI call has completed.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCI_Seek(ByVal DeviceAlias As String, ByVal SeekPosistion As Long, Optional ByVal hCallbackProc As Long, Optional ByVal WaitToComplete As Boolean = True, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  ' Make sure parameters are valid
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  
  ' Seek to the specified position of the file
  MCI_Seek = MCIa_SendCommand("seek " & DeviceAlias & " to " & CStr(SeekPosistion), , , hCallbackProc, WaitToComplete, ShowErrors)
  
End Function

'=============================================================================================================
' MCI_SeekBeginning
'
' Purpose :
' Moves the current play position to the beginning of the media.
'
' Param                Use
' ------------------------------------
' DeviceAlias          Specifies the name of the device to perform the action on.
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' WaitToComplete       Optional. If this is set to TRUE, this function will not return control to the application
'                      until the MCI call has completed.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCI_SeekBeginning(ByVal DeviceAlias As String, Optional ByVal hCallbackProc As Long, Optional ByVal WaitToComplete As Boolean = True, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  ' Make sure parameters are valid
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  
  ' Seek to the beginning of the file
  MCI_SeekBeginning = MCIa_SendCommand("seek " & DeviceAlias & " to start", , , hCallbackProc, WaitToComplete, ShowErrors)
  
End Function

'=============================================================================================================
' MCI_SeekEnd
'
' Purpose :
' Moves the current play position to the end of the media.
'
' Param                Use
' ------------------------------------
' DeviceAlias          Specifies the name of the device to perform the action on.
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' WaitToComplete       Optional. If this is set to TRUE, this function will not return control to the application
'                      until the MCI call has completed.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCI_SeekEnd(ByVal DeviceAlias As String, Optional ByVal hCallbackProc As Long, Optional ByVal WaitToComplete As Boolean = True, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  ' Make sure parameters are valid
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  
  ' Seek to the end of the file
  MCI_SeekEnd = MCIa_SendCommand("seek " & DeviceAlias & " to end", , , hCallbackProc, WaitToComplete, ShowErrors)
  
End Function


'=============================================================================================================
' MCI_Set
'
' Purpose :
' Alters the specified device or device playback by sending the specified command to it.  The commands vary
' between device types and not all commands are compatible with all device types.  See table below for details.
'
' *NOTE - This function makes the application wait for the device parameter to be set because it must finish
'         being set before further calls to it can be made successfully
'
' Param                Use
' ------------------------------------
' DeviceAlias          Specifies the name of the device to perform the action on.
' ControlSettingString Specifies the string to pass to set the device condition or information. This can be
'                      any of the applicable commands below and can do a wide variety of things to the device.
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'-------------------------------------------------------------------------------------------------------------
'
' Recognized device setting strings:
'=============================================================================================================
'ControlSettingString:        Applies To: Meaning:
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'alignment <integer>          WaveAudio   Sets the alignment of data blocks relative to the start of data passed to the waveform-audio device. The file is saved in this format.
'any input                    WaveAudio   Use any input that supports the current format when recording. This is the default setting.
'any output                   WaveAudio   Use any output that supports the current format when playing. This is the default.
'assemble record on/off       VCR         In assemble mode, all tracks are recorded as defined by the device. Most VCRs are in assemble mode by default.
'audio all on/off             AVIVideo,   Disables or enables audio output. Video-overlay devices, the MCISEQ sequencer, and the MCIWAVE waveform-audio device do not support this flag.
'                             CDAudio,
'                             Overlay,
'                             Sequencer,
'                             VCR,
'                             VideoDisc,
'                             WaveAudio
'audio left on/off            AVIVideo,   Disables or enables output to either the left or the right audio channel. Video-overlay devices, the MCISEQ sequencer, and the MCIWAVE waveform-audio device do not support this flag.
'                             CDAudio,
'                             Overlay,
'                             Sequencer,
'                             VCR,
'                             VideoDisc,
'                             WaveAudio
'audio right on/off           AVIVideo,   Disables or enables output to either the left or the right audio channel. Video-overlay devices, the MCISEQ sequencer, and the MCIWAVE waveform-audio device do not support this flag.
'                             CDAudio,
'                             Overlay,
'                             Sequencer,
'                             VCR,
'                             VideoDisc,
'                             WaveAudio
'bitspersample <bit_count>    WaveAudio   Sets the number of bits per PCM (Pulse Code Modulation) sample played or recorded. The file is saved in this format.
'bytespersec <byte_rate>      WaveAudio   Sets the average number of bytes per second played or recorded. The file is saved in this format.
'channels <channel_count>     WaveAudio   Sets the channels for playing and recording. The file is saved in this format.
'clock <time>                 VCR         Sets time on the external clock to time. The format is specified as a long unsigned integer.
'counter format               VCR         Set the time format for the counter, as returned by status "counter". For information about applicable types, see the set "time format" command.
'counter <value>              VCR         Sets the VCR counter to the specified value. The value must be specified in the current counter format. For more information, see the set "counter format" command.
'door closed                  AVIVideo,   Retracts the tray and closes the door, if possible. For VCRs, loads the tape automatically.
'                             CDAudio,
'                             Overlay,
'                             Sequencer,
'                             VCR,
'                             VideoDisc,
'                             WaveAudio
'door open                    AVIVideo,   Opens the door and ejects the tray or tape, if possible.
'                             CDAudio,
'                             Overlay,
'                             Sequencer,
'                             VCR,
'                             VideoDisc,
'                             WaveAudio
'file format <format>         AVIVideo    Specifies a file format that is used for save or capture commands. If omitted, this might default to a device driver defined format. If the specified file format conflicts with the currently selected algorithm and quality, then they are changed to the defaults for the file format. The following file formats are defined:
'                                           AVI   - Specifies AVI format
'                                           AVSS  - Specifies AVSS format
'                                           DIB   - Specifies DIB format
'                                           JFIF  - Specifies JFIF format
'                                           JPEG  - Specifies JPEG format
'                                           MPEG  - Specifies MPEG format
'                                           RDIB  - Specifies RLE DIB format
'                                           RJPEG - Specifies RJPEG format
'format tag pcm               WaveAudio   Sets the format type to PCM for playing and recording. The file is saved in this format.
'format tag <tag>             WaveAudio   Sets the format type for playing and recording. The file is saved in this format.
'Index counter                VCR         (see "index timecode")
'Index date                   VCR         (see "index timecode")
'index time                   VCR         (see "index timecode")
'Index timecode               VCR         Sets the current display screen on the VCR
'input <integer>              WaveAudio   Sets the audio channel used as the input
'length <duration>            VCR         Sets the user-specified length of the tape in the VCR. This length is returned by the status "length" command and is provided for compatibility with applications that require this command to return a valid length.
'master midi                  Sequencer   Sets the MIDI sequencer as the synchronization source. Synchronization data is sent in MIDI format. The MCISEQ sequencer does not support this flag.
'master none                  Sequencer   Inhibits the MIDI sequencer from sending synchronization data. The MCISEQ sequencer does not support this flag.
'master smpte                 Sequencer   Sets the MIDI sequencer as the synchronization source. Synchronization data is sent in SMPTE (Society of Motion Picture and Television Engineers) format. The MCISEQ sequencer does not support this flag.
'offset <time>                Sequencer   Sets the SMPTE offset time. The offset is the beginning time of a SMPTE based sequence. The time is expressed as hh:mm:ss:ff, where hh is hours, mm is minutes, ss is seconds, and ff is frames.
'output <integer>             WaveAudio   Sets the audio channel used as the output.
'pause <timeout>              VCR         Sets the maximum duration, in milliseconds, of a pause command. A timeout value of zero indicates that no time-out will occur.
'postroll duration <duration> VCR         Sets the length, in the current time format, needed to brake the VCR transport when a stop or pause command is issued.
'port mapper                  Sequencer   Sets the MIDI mapper as the port receiving the MIDI messages. This command fails if the MIDI mapper or a port it needs is being used by another application.
'port none                    Sequencer   Disables the sending of MIDI messages. This command also closes a MIDI port.
'port <port_number>           Sequencer   Sets the MIDI port receiving the MIDI messages. This command fails if the port you are trying to open is being used by another application.
'power on/off                 VCR         Sets the device power to on or off.
'preroll duration <duration>  VCR         Sets the length, in the current time format, needed to stabilize the VCR output.
'record format SP/EP/LP       VCR         Sets the recording mode for the VCR to SP for standard play, EP for extended play, or LP for long play. These values are not intended to be VHS specific. They map to any three appropriate modes with other tape formats. For example, SP maps to the fastest, highest quality recording.
'samplespersec <integer>      WaveAudio   Sets the sample rate for playing and recording. The file is saved in this format.
'seek exactly on/off          AVIVideo    Selects one of two seek modes. With "seek exactly on", seek will always move to the frame specified. With "seek exactly off", seek will move to the closest key frame prior to the frame specified.
'slave file                   Sequencer   Sets the MIDI sequencer to use file data as the synchronization source. This is the default setting.
'slave midi                   Sequencer   Sets the MIDI sequencer to use incoming MIDI data for the synchronization source. The sequencer recognizes synchronization data with the MIDI format. The MCISEQ sequencer does not support this flag.
'slave none                   Sequencer   Sets the MIDI sequencer to ignore synchronization data.
'slave smpte                  Sequencer   Sets the MIDI sequencer to use incoming MIDI data for the synchronization source. The sequencer recognizes synchronization data with the SMPTE format. The MCISEQ sequencer does not support this flag.
'speed <factor>               AVIVideo,   Sets the relative speed of video and audio playback from the workspace. Factor is the ratio between the nominal frame rate and the desired frame rate, where the nominal frame rate is designated as 1000. (A rate of 500 is half normal speed, 2000 is twice normal speed, and so on.) Setting the speed to zero plays the video as fast as possible without dropping frames and without audio.
'                             VCR
'still file format <format>   AVIVideo    Specifies the file format used for capture commands.
'tempo <tempo_value>          Sequencer   Sets the tempo of the sequence according to the current time format. For a PPQN-based file, the tempo_value is interpreted as beats per minute. For a SMPTE-based file, the tempo_value is interpreted as frames per second.
'time format bytes            WaveAudio   In a PCM file format, sets the time format to bytes. All position information is specified as bytes following this command.
'time format frames           AVIVideo,   Sets the time format to frames. All commands that use position values will assume frames. When the device is opened, frames is the default mode. Supported by videodiscs using CAV format.
'                             VCR,
'                             VideoDisc
'time format hms              VCR,        Sets the time format to hours, minutes, and seconds. All commands that use position values will assume HMS. HMS is the default format for CLV videodiscs.  Specify an HMS value as hh:mm:ss, where hh is hours, mm is minutes, and ss is seconds. You can omit a field if it and all following fields are zero  (for example - 3, 3:0, and 3:0:0 are all valid ways to express 3 hours).
'                             VideoDisc
'time format milliseconds     AVIVideo,   Sets the time format to milliseconds. All commands that use position values will assume milliseconds. You can abbreviate milliseconds as "ms".  For sequencer devices, the sequence file sets the default format to PPQN or SMPTE. Video-overlay devices do not support this flag.
'                             CDAudio,
'                             Sequencer,
'                             VCR,
'                             VideoDisc,
'                             WaveAudio
'time format msf              CDAudio,    Sets the time format to minutes, seconds, and frames. All commands that use position values will assume MSF (the default format for CD audio).  Specify an MSF value as mm:ss:ff, where mm is minutes, ss is seconds, and ff is frames. You can omit a field if it and all following fields are zero. For example, 3, 3:0, and 3:0:0 are valid ways to express 3 minutes.  The MSF fields have the following maximum values: Minutes 99, Seconds 59, Frames 74
'                             VCR
'time format samples          WaveAudio   Sets the time format to samples. All position information is specified as samples following this command.
'time format smpte 24         Sequencer,  (see "time format smpte 30")
'                             VCR
'time format smpte 25         Sequencer,  (see "time format smpte 30")
'                             VCR
'time format smpte 30         Sequencer,  Sets the time format to an SMPTE frame rate.  For VCRs, sets the time format to hh:mm:ss:ff, where the legal values are 00:00:00:00 through 23:59:59:xx, and xx is one less than the frames per second as specified by the number 24, 25, or 30 as specified in the flag. On input, colons (:) are required to separate the components. The least significant units can be omitted if they are 00; for example, 02:00 is the same as 02:00:00:00.  All commands that use position values will assume SMPTE format.  The sequence file sets the default format to PPQN or SMPTE.
'                             VCR
'time format smpte 30 drop    Sequencer,  Sets the time format to SMPTE 30 drop frame rate.  For VCRs, same as SMPTE 30, except that certain timecode positions are dropped from the format to have the recorded timecode positions for each frame (at the NTSC frame rate of 29.97 fps) correspond to real time (at 30 fps). Timecode positions that are dropped are as follows: two every minute, on the minute, for the first nine of every ten minutes of recorded content. For example, at 01:04:59:29, the next timecode position would be 01:05:00:02, not 01:05:00:00.  All commands that use position values will assume SMPTE format.  The sequence file sets the default format to PPQN or SMPTE.
'                             VCR
'time format song pointer     Sequencer   Sets the time format to song pointer (sixteenth notes). All commands that use position values will assume song pointer units. This flag is valid only for a sequence of division type PPQN.
'time format tmsf             CDAudio     Sets the time format to tracks, minutes, seconds, and frames. All commands that use position values will assume TMSF.  Specify a TMSF value as tt:mm:ss:ff, where tt is tracks, mm is minutes, ss is seconds, and ff is frames. You can omit a field if it and all following fields are zero. For example, 3, 3:0, 3:0:0, and 3:0:0:0 are all valid ways to express track 3.  The TMSF fields have the following maximum values: Tracks 99, Minutes 99, Seconds 59, Frames 74
'time format track            VCR,        Sets the position format to tracks. All commands that use position values will assume tracks.
'                             VideoDisc
'time mode counter            VCR         Sets the position-information mode to use the VCR counters.
'time mode detect             VCR         Sets the position information mode based on detection of timecode information on the tape. If timecode information is detected, the time type is set to "timecode"; otherwise, the time type is set to "counter".  "Detect" is a special mode. Whenever the driver is opened, a new tape is inserted, or the "time mode" command is issued, the driver checks the current time mode available on the tape and sets "time type" to either "timecode" or "counter". Once "time type" is set, the driver doesn't change it until one of the above conditions occurs again.
'time mode timecode           VCR         Sets the position information mode to use "timecode" information on the tape.
'tracking plus                VCR         (see "tracking reset")
'tracking minus               VCR         (see "tracking reset")
'tracking reset               VCR         Adjusts the speed of the videotape transport in fine increments. Use the "tracking" flags when a noisy picture is obtained from a VCR. "Tracking plus" increases the transport speed. "Tracking minus" decreases the transport speed. "Tracking reset" returns the tracking adjustment to zero.
'video on/off                 AVIVideo,   Enables or disables video output
'                             Overlay,
'                             VideoDisc
'
'=============================================================================================================
Public Function MCI_Set(ByVal DeviceAlias As String, ByVal ControlSettingString As String, Optional ByVal hCallbackProc As Long, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  ' Make sure parameters are valid
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  
  ' Set the specified parameter for the device
  MCI_Set = MCIa_SendCommand("set " & DeviceAlias & " " & ControlSettingString, , , hCallbackProc, True, ShowErrors)
  
End Function


'=============================================================================================================
' MCI_Status
'
' Purpose :
' Retrieves the specified information about the AVI device or AVI playback by sending the specified status
' request command to it.  Not all commands are compatible with all multimedia systems.  See table below for details.
'
' *NOTE - This function makes the application wait for the device status to be retrieved because it must finish
'         getting the status before the information returned by it can be used successfully
'
' Param                Use
' ------------------------------------
' DeviceAlias          Specifies the name of the device to perform the action on.
' StatusRequestString  Specifies the status type being requested in the form of a string.  The status can be
'                      any of the statuses in the following list
' Return_Status        Returns the status requested in the "StatusRequestString" parameter in the form of a string.
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'-------------------------------------------------------------------------------------------------------------
'
' Recognized device status request strings:
'=============================================================================================================
'StatusRequestString:        Applies To: Meaning:
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'alignment                   WaveAudio   Returns the block alignment of data, in bytes.
'assemble record             VCR         Returns TRUE if the device is set to assemble mode recording.
'audio                       AVIVideo    Returns "on" or "off" depending on the most recent setaudio "on" or "off" command. It returns "on" if either or both speakers are enabled, and "off" otherwise.
'audio alignment             AVIVideo    Returns the alignment of data blocks relative to the start of input waveform-audio data.
'audio bitspersample         AVIVideo    Returns the number of bits per sample the device uses for recording. This flag applies only to devices supporting the "pcm" algorithm.
'audio breaks                AVIVideo    Returns the number of times the audio portion of the last AVI sequence broke up. The system counts an audio break whenever it attempts to write audio data to the device driver and discovers that the driver has already played all of the available data. This flag is recognized only by the MCIAVI digital-video driver. It is meant for performance evaluation only; the return value is difficult to interpret.
'audio bytespersec           AVIVideo    Returns the average number of bytes per second used for recording.
'audio input                 AVIVideo    Returns the approximate instantaneous audio level of the analog input audio signal. A value greater than 1000 implies clipping distortion. Some devices can return this value only while recording audio. The value has no associated set or setaudio command.
'audio monitor               VCR         Returns "output", or one of the valid source-input types. For more information, see the setaudio "monitor" command.
'audio monitor number        VCR         Returns the monitored-video number of the type specified by status "audio monitor". For more information, see the setaudio command.
'audio record                AVIVideo,   Returns "on" or "off", reflecting the state set by setaudio "record".
'                            VCR
'audio record track <number> VCR         Returns TRUE if the VCR is set to record audio. If no track number is given, the default value of 1 is assumed.
'audio samplespersec         AVIVideo    Returns the number of samples per second recorded.
'audio source                AVIVideo,   Returns the current audio digitizer source: "left", "right", "average", or "stereo".
'                            VCR
'audio source number         VCR         Returns the audio-source number of the type returned by status "audio source". For more information, see the setaudio command.
'audio stream                AVIVideo    Returns the current audio-stream number.
'bass                        AVIVideo    Returns the current audio-bass level.
'bitsperpel                  AVIVideo    Returns the number of bits per pixel for saving captured or recorded data.
'bitspersample               WaveAudio   Returns the bits per sample.
'brightness                  AVIVideo    Returns the current video-brightness level.
'bytespersec                 WaveAudio   Returns the average number of bytes per second played or recorded.
'cdaudio type track <number> CDAudio     Returns the type of the specified track number. This can be "audio" or "other."
'channel                     VCR         Returns the integer value of the channel set on the tuner.
'channel tuner number        VCR         If "tuner" number is given, then the currently selected channel on the logical tuner number will be returned. Note that there can be several logical tuners.
'channels                    WaveAudio   Returns the number of channels set (1 for mono, 2 for stereo).
'clock                       VCR         Returns the external time. The time must be an unsigned long integer expressing total increments. For more information, see the capability "clock increment rate" command.
'clock id                    VCR         Returns a unique integer identifying the clock.
'color                       AVIVideo    Returns the current color level.
'contrast                    AVIVideo    Returns the current contrast level.
'counter                     VCR         Returns the counter position, in the current counter format.
'counter format              VCR         Returns the current counter format. For more information, see the set "counter format" command.
'counter resolution          VCR         Returns "frames" or "seconds", indicating the counter's resolution. This is not the same as accuracy.
'current track               AVIVideo,   Returns the current track. The MCISEQ sequencer returns 1.
'                            CDAudio,
'                            Sequencer,
'                            VCR,
'                            VideoDisc,
'                            WaveAudio
'disc size                   VideoDisc   Returns either 8 or 12, indicating the size of the loaded disc in inches.
'disk space <drive>          AVIVideo,   Returns the approximate disk space, in the current time format, that can be obtained by a reserve command for the specified disk drive. The drive is usually specified as a single letter or a single letter followed by a colon (:). Some devices, however, might use a path.
'division type               Sequencer   Use this information to determine the format of the MIDI file and the meaning of tempo and position information.  Returns one of the following file division types: PPQN, SMPTE 24 frame, SMPTE 25 frame, SMPTE 30 drop frame, SMPTE 30 frame
'file completion             AVIVideo    Returns the estimated percentage a load, save, capture, cut, copy, delete, paste, or undo operation has progressed. (Applications can use this to provide a visual indicator of progress.)
'file format                 AVIVideo    Returns the current file format for record or save commands.
'file mode                   AVIVideo    Returns "loading", "saving", "editing", or "idle". During a load operation, it returns "loading". During save and capture operations, it returns "saving". During cut, copy, delete, paste, or undo operations, it returns "editing".
'format tag                  WaveAudio   Returns the format tag.
'forward                     AVIVideo,   Returns TRUE if the play direction is forward or if the device is not playing.
'                            VideoDisc
'frame rate                  VCR         Returns the number of frames per second that the device will use by default. NTSC devices return 30, PAL 25, and so on.
'frames skipped              AVIVideo    Returns the number of frames that were not drawn when the last AVI sequence was played. This flag is recognized only by the MCIAVI digital-video driver. It is meant for performance evaluation only; the return value is difficult to interpret.
'gamma                       AVIVideo    Returns the value set with setvideo "gamma to" value.
'index                       VCR         Returns the current index display. For more information, see the set "index" command.
'index on                    VCR         Returns TRUE if the index is on.
'input                       AVIVideo    Returns the input set. If one is not set, the error returned indicates that any device can be used.  For AVIVideo devices, modifies the "bass", "treble", "volume", "brightness", "color", "contrast", "gamma", "sharpness", or "tint" flag so that it applies only to the input. This is the default when monitoring the input.
'                            WaveAudio
'left volume                 AVIVideo    Returns the volume set for the left audio channel.
'length                      AVIVideo,   Returns the total length of the media, in the current time format.  For PPQN files, the length is returned in song pointer units. For SMPTE files, it is returned as hh:mm:ss:ff, where hh is hours, mm is minutes, ss is seconds, and ff is frames. For VCR devices, the length is 2 hours (unless the length has been explicitly changed using the set command).
'                            CDAudio,
'                            Sequencer,
'                            VCR,
'                            VideoDisc,
'                            WaveAudio
'length track <number>       AVIVideo,   Returns the length of the track, in time or frames, specified by number.  For PPQN files, the length is returned in song pointer units. For SMPTE files, it is returned as hh:mm:ss:ff, where hh is hours, mm is minutes, ss is seconds, and ff is frames.
'                            CDAudio,
'                            Sequencer,
'                            VCR,
'                            VideoDisc,
'                            WaveAudio
'level                       WaveAudio   Returns the current PCM audio sample value.
'master                      Sequencer   Returns "midi", "none", or "smpte" depending on the type of synchronization set.
'media present               AVIVideo,   Returns TRUE if the media is inserted in the device or FALSE otherwise. Sequencer, video-overlay, digital-video, and waveform-audio devices return TRUE.
'                            CDAudio,
'                            Overlay,
'                            Sequencer,
'                            VCR,
'                            VideoDisc,
'                            WaveAudio
'media type                  VCR         Returns the type of the media. For VCRS, this is "8mm", "vhs", "svhs", "beta", "Hi8", "edbeta", or "other".  For videodiscs, this is "CAV", "CLV", or "other", depending on the type of videodisc.
'                            VideoDisc
'mode                        AVIVideo,   Returns the current mode of the device. All devices can return the "not ready", "paused", "playing", and "stopped" values. Some devices can return the additional "open", "parked", "recording", and "seeking" values.
'                            CDAudio,
'                            Overlay,
'                            Sequencer,
'                            VCR,
'                            VideoDisc,
'                            WaveAudio
'monitor                     AVIVideo    Returns "file" or "input". The returned value indicates the current presentation source.
'monitor method              AVIVideo    Returns "pre", "post", or "direct". The returned value indicates the method used for input monitoring.
'nominal                     AVIVideo    The item modifies the "bass", "brightness", "color", "contrast", "gamma", "sharpness", "tint", "treble," and "volume" flags to return the nominal value instead of the current setting.
'nominal frame rate          AVIVideo    Returns the nominal frame rate associated with the file. The units are in frames per second multiplied by 1000.
'nominal record frame rate   AVIVideo    Returns the nominal frame rate associated with the input video signal. The units are in frames per second multiplied by 1000.
'number of audio tracks      VCR         Returns the number of audio tracks on the media.
'number of tracks            AVIVideo,   Returns the number of tracks on the media. The MCISEQ and MCIWAVE devices return 1, as do most VCR devices. The MCIPIONR device does not support this flag.
'                            CDAudio,
'                            Overlay,
'                            Sequencer,
'                            VCR,
'                            VideoDisc,
'                            WaveAudio
'number of video tracks      VCR         Returns the number of video tracks on the media.
'offset                      Sequencer   Returns the offset of a SMPTE-based file. The offset is the start time of a SMPTE-based sequence. The time is returned as hh:mm:ss:ff, where hh is hours, mm is minutes, ss is seconds, and ff is frames.
'output                      AVIVideo,   Returns the currently set output. If no output is set, the error returned indicates that any device can be used.  For digital-video devices, modifies the "bass", "treble", "volume", "brightness", "color", "contrast", "gamma", "sharpness", or "tint" flag so that it applies only to the output. This is the default when monitoring a file.
'                            WaveAudio
'pause mode                  AVIVideo    Returns "recording" if the device is paused while recording. It returns "playing" if the device is paused while playing. It returns the error "Action not applicable in current mode" if the device is not paused.
'pause <timeout>             VCR         Returns the maximum duration, in milliseconds, of a pause command.
'play format                 VCR         Returns a code indicating the format that the videotape will be played back in, if detectable: "lp", "ep", "sp", or "other". For more information, see the "record format" flag.
'play speed                  AVIVideo    Returns a value representing how closely the actual playing time of the last AVI sequence matched the target playing time. The value 1000 indicates that the target time and the actual time were the same. A value of 2000, for example, would indicate that the AVI sequence took twice as long to play as it should have. This flag is recognized only by the MCIAVI digital-video driver. It is meant for performance evaluation only; the return value is difficult to interpret.
'port                        Sequencer   Returns the MIDI port number assigned to the sequence.
'position                    AVIVideo,   Returns the current position.  For PPQN files, the position is returned in song pointer units. For SMPTE files, it is returned as hh:mm:ss:ff, where hh is hours, mm is minutes, ss is seconds, and ff is frames.
'                            CDAudio,
'                            Sequencer,
'                            VCR,
'                            VideoDisc,
'                            WaveAudio
'position start              VCR         Returns the position of the start of the usable media.
'position track <number>     AVIVideo,   Returns the position of the start of the track specified by number.  For PPQN files, the time format is returned in song pointer units. For SMPTE files, it is returned as hh:mm:ss:ff, where hh is hours, mm is minutes, ss is seconds, and ff is frames. The MCISEQ sequencer returns zero. The MCIPIONR device does not support this flag. The MCIWAVE device returns zero.
'                            CDAudio,
'                            Sequencer,
'                            VCR,
'                            VideoDisc,
'                            WaveAudio
'postroll duration           VCR         Returns the length of videotape, in the current time format, needed to brake the VCR transport when a stop or pause command is issued.
'power on                    VCR         Returns TRUE if the VCR's power is on.
'preroll duration            VCR         Returns the length of videotape, in the current time format, needed to stabilize the VCR output.
'ready                       AVIVideo,   Returns TRUE if the device is ready to accept another command.
'                            CDAudio,
'                            Overlay,
'                            Sequencer,
'                            VCR,
'                            VideoDisc,
'                            WaveAudio
'record format               VCR         Returns a code indicating the format that the videotape will be recorded in. The current return types are "lp", "ep", "sp", or "other". These formats are not VHS specific and can be applied to any VCR that has multiple selectable recording formats. The "sp" type is the fastest, highest quality recording format and is used as default on single format VCRs.
'record frame rate           AVIVideo    Returns the frame rate, in frames per second multiplied by 1000, used for compression.
'reference frame             AVIVideo    Returns the frame number for the nearest key frame image that precedes the specified frame.
'reserved size               AVIVideo    Returns the size, in the current time format, of the reserved workspace. The size corresponds to the approximate time it would take to play the compressed data from a full workspace. It returns zero if there is no reserved disk space. This flag returns the approximate size because the precise disk space for compressed data cannot, in general, be predicted until after the data has been compressed.
'right volume                AVIVideo    Returns the volume set for the right audio channel.
'samplespersec               WaveAudio   Returns the number of samples per second played or recorded.
'seek exactly                AVIVideo    Returns "on" or "off", indicating whether or not the "seek exactly" flag is set.
'sharpness                   AVIVideo    Returns the current sharpness level of the device.
'side                        VideoDisc   Returns 1 or 2 to indicate which side of the videodisc is loaded.
'slave                       Sequencer   Returns "file", "midi", "none", or "smpte" depending on the type of synchronization set.
'smpte                       AVIVideo    Returns the SMPTE timecode associated with the current position in the workspace. This is a string with the form dd:dd:dd:dd, where each d denotes a digit from 0 to 9. If the workspace data does not include timecode data, then this flag returns 00:00:00:00.
'speed                       AVIVideo,   Returns the current speed of the device in frames per second (or in the same format used by the set "speed" command). The MCIPIONR videodisc player does not support this flag.
'                            VCR,
'                            VideoDisc
'start position              AVIVideo,   Returns the starting position of the media.
'                            CDAudio,
'                            Sequencer,
'                            VideoDisc,
'                            WaveAudio
'still file format           AVIVideo,   Returns the current file format for the capture command.
'stretch                     Overlay     Returns TRUE if stretching is enabled.
'tempo                       Sequencer   Returns the current tempo of a MIDI sequence in the current time format. For files with PPQN format, the tempo is in beats per minute. For files with SMPTE format, the tempo is in frames per second.
'time format                 AVIVideo,   Returns the current time format. For more information, see the time formats in the set command.
'                            CDAudio,
'                            Overlay,
'                            Sequencer,
'                            VCR,
'                            VideoDisc,
'                            WaveAudio
'time mode                   VCR         Returns the current position time mode. It can be "detect", "timecode", or "counter".
'time type                   VCR         Returns the current position time in use: "timecode" or "counter".
'timecode present            VCR         Returns TRUE if timecode has been recorded at the current position on the tape. The timecode must advance from the current position. A VCR might need to be played to check this condition.
'timecode record             VCR         Returns TRUE if the VCR is set to record timecode.
'timecode type               VCR         Returns "smpte", "smpte drop", "other", or "none". Note the frames per second can be obtained from the status "frame rate" command, and the accuracy of the device can be returned by the capability "seek accuracy" command.
'tint                        AVIVideo    Returns the current video-tint level.
'treble                      AVIVideo    Returns the current audio-treble level.
'tuner number                VCR         Returns the current logical-tuner number.
'unsaved                     AVIVideo    Returns TRUE if there is recorded data in the workspace that might be lost as a result of a close, load, record, reserve, cut, delete, or paste command. Returns FALSE otherwise.
'video                       AVIVideo    Returns "on" or "off", reflecting the state set by the setvideo command.
'video key color             AVIVideo    Returns the value for the key color.
'video key index             AVIVideo    Returns the value for the key index.
'video monitor               VCR         Returns "output" or one of the valid source-input types. For more information, see the setvideo "monitor" command.
'video monitor number        VCR         Returns the monitored-video number of the type returned by status "video monitor". For more information, see the setvideo command.
'video record                AVIVideo,   Returns "on" or "off", reflecting the current state set by setvideo "record".
'                            VCR
'video record track <number> VCR         Return TRUE if the VCR is set to record video. If no track number is given, the default value of 1 is assumed.
'video source                AVIVideo,   Returns the video-source type. For more information, see the setvideo command.
'                            VCR
'video source number         AVIVideo,   Returns a number corresponding to the video source of the type in use. For example, it returns 2 if the second NTSC video source input is being used.
'                            VCR
'video stream                AVIVideo    Returns the current video-stream number.
'volume                      AVIVideo    Returns the average volume to the left and right speaker. This returns an error if the device has not been played or volume has not been set.
'window handle               AVIVideo,   Returns ASCII decimal value for the window handle in the low-order word of the return value.
'                            Overlay
'window maximized            AVIVideo    Returns TRUE if the window is maximized.
'window minimized            AVIVideo    Returns TRUE if the window is minimized.
'window visible              AVIVideo    Returns TRUE if the window is not hidden.
'write protected             VCR         Returns TRUE if the device detects that it cannot record (that is, if the write protect is on). If it can record, or if it is unable to determine whether or not it can record (without actually writing), the driver returns FALSE.
'
'=============================================================================================================
Public Function MCI_Status(ByVal DeviceAlias As String, ByVal StatusRequestString As String, ByRef Return_Status As String, Optional ByVal hCallbackProc As Long, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  ' Clear the return variable
  Return_Status = ""
  
  ' Make sure parameters are valid
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  
  ' Get the specified status for the specified device
  MCI_Status = MCIa_SendCommand("status " & DeviceAlias & " " & StatusRequestString, True, Return_Status, hCallbackProc, True, ShowErrors)
  
End Function

'=============================================================================================================
' MCI_Stop
'
' Purpose :
' Stops the playback of the specified device.  Note that this does NOT reset the current position in the
' media being played to the beginning.  To do that, make a call to the "MCI_SeekBeginning" method.
'
' Param                Use
' ------------------------------------
' DeviceAlias          Specifies the name of the device to perform the action on.
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' WaitToComplete       Optional. If this is set to TRUE, this function will not return control to the application
'                      until the MCI call has completed.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCI_Stop(ByVal DeviceAlias As String, Optional ByVal hCallbackProc As Long, Optional ByVal WaitToComplete As Boolean = False, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  ' Make sure parameters are valid
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  
  ' Stop the device
  MCI_Stop = MCIa_SendCommand("stop " & DeviceAlias, , , hCallbackProc, WaitToComplete, ShowErrors)
  
End Function

'=============================================================================================================
' MCI_WindowCaption
'
' Purpose :
' This function changes the caption of the window that is displaying any type of VIDEO playback.
'
' Param                Use
' ------------------------------------
' DeviceAlias          Specifies the name of the device to perform the action on.
' strNewWindowCaption  Specifies the new window caption to be set.
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' WaitToComplete       Optional. If this is set to TRUE, this function will not return control to the application
'                      until the MCI call has completed.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCI_WindowCaption(ByVal DeviceAlias As String, ByVal strNewWindowCaption As String, Optional ByVal hCallbackProc As Long, Optional ByVal WaitToComplete As Boolean = False, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  Dim hWindow   As Long
  Dim strReturn As String
  
  ' Make sure parameters are valid
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  
  ' Get the window handle, and if valid... use to change the caption
  If MCI_Status(DeviceAlias, "window handle", strReturn, hCallbackProc, ShowErrors) = True Then
    If strReturn <> "" Then
      If IsNumeric(strReturn) = True Then
        hWindow = CLng(strReturn)
        If SetWindowText(hWindow, strNewWindowCaption & Chr(0)) <> 0 Then MCI_WindowCaption = True
      End If
    End If
  End If
  
End Function

'=============================================================================================================
' MCI_WindowIcon
'
' Purpose :
' This function changes the icon of the window that is displaying any type of VIDEO playback.
'
' Param                Use
' ------------------------------------
' DeviceAlias          Specifies the name of the device to perform the action on.
' hIcon                Specifies the handle of the icon to change the window's icon to.
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' WaitToComplete       Optional. If this is set to TRUE, this function will not return control to the application
'                      until the MCI call has completed.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCI_WindowIcon(ByVal DeviceAlias As String, ByVal hIcon As Long, Optional ByVal hCallbackProc As Long, Optional ByVal WaitToComplete As Boolean = False, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  Dim hWindow   As Long
  Dim strReturn As String
  
  ' Make sure parameters are valid
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  
  ' Get the window handle, and if valid... use to change the caption
  If MCI_Status(DeviceAlias, "window handle", strReturn, hCallbackProc, ShowErrors) = True Then
    If strReturn <> "" Then
      If IsNumeric(strReturn) = True Then
        hWindow = CLng(strReturn)
        SendMessage lngWindowHandle, WM_SETICON, ICON_BIG, ByVal hIcon
      End If
    End If
  End If
  
End Function



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



'=============================================================================================================
' MCIa_MoveVideo
'
' Purpose :
' This method allows you to manipulate the playback of video device types.  You can use this method to move
' the location of the popup window playing the video, to modify the source coordinates of the video to give
' the effect of zooming in, or to modify the destination coordinates of the video to give the effect of
' zoming out.
'
' Param                      Use
' ------------------------------------
' DeviceAlias                Specifies the name of the device to perform the action on.
' MoveWindowLocation         Optional. If set to TRUE, the popup window will be moved according to the next 4 params
' Window_Left                Optional. Specifies the coordinates of the left edge of the popup window after it's moved
' Window_Top                 Optional. Specifies the coordinates of the top edge of the popup window after it's moved
' Window_Width               Optional. Specifies the new width of the popup window after it's moved
' Window_Height              Optional. Specifies the new height of the popup window after it's moved
' DisplayVideoOnEntireWindow Optional. If set to TRUE, the entire window is used to display the video media.
'                            Otherwise the following 4 params are used to specify where on the window to display the
'                            video.
' Dest_Left                  Optional. Specifies the coordinates of the left edge of the video display location
' Dest_Top                   Optional. Specifies the coordinates of the top edge of the video display location
' Dest_Width                 Optional. Specifies the new width of the video display area
' Dest_Height                Optional. Specifies the new height of the video display area
' UseEntireSourceImage       Optional. If set to TRUE, the entire source image is used to display the video media.
'                            Otherwise the following 4 params are used to specify the source of the display.
' Source_Left                Optional. Specifies the coordinates of the left edge of the source of the display
' Source_Top                 Optional. Specifies the coordinates of the top edge of the source of the display
' Source_Width               Optional. Specifies the new width of the source of the display
' Source_Height              Optional. Specifies the new height of the source of the display
' hCallbackProc              Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                            the device completes an action.  In order for this to work, you must specify the
'                            handle to the window in your initial call to MCI_Open, then specify the handle of
'                            that window in subsequent calls to methods like this.
' WaitToComplete             Optional. If this is set to TRUE, this function will not return control to the application
'                            until the MCI call has completed.
' ShowErrors                 Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                            the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCIa_MoveVideo(ByVal DeviceAlias As String, Optional ByVal MoveWindowLocation As Boolean = False, Optional ByVal Window_Left As Long, Optional ByVal Window_Top As Long, Optional ByVal Window_Width As Long, Optional ByVal Window_Height As Long, Optional ByVal DisplayVideoOnEntireWindow As Boolean = True, Optional ByVal Dest_Left As Long, Optional ByVal Dest_Top As Long, Optional ByVal Dest_Width As Long, Optional ByVal Dest_Height As Long, Optional ByVal UseEntireSourceImage As Boolean = True, Optional ByVal Source_Left As Long, Optional ByVal Source_Top As Long, Optional ByVal Source_Width As Long, Optional ByVal Source_Height As Long, Optional ByVal hCallbackProc As Long, Optional ByVal WaitToComplete As Boolean = False, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  ' Make sure parameters are valid
  DeviceAlias = Trim(DeviceAlias)
  If DeviceAlias = "" Then Exit Function
  
  ' Restores the original window/client and displays if not already displayed
  If MCIa_SendCommand("put " & DeviceAlias & " window", , , hCallbackProc, WaitToComplete, ShowErrors) = False Then Exit Function
  If MCIa_SendCommand("put " & DeviceAlias & " window client", , , hCallbackProc, WaitToComplete, ShowErrors) = False Then Exit Function
  
  ' Set the source of the image
  If UseEntireSourceImage = True Then
    If MCIa_SendCommand("put " & DeviceAlias & " source", , , hCallbackProc, WaitToComplete, ShowErrors) = False Then Exit Function
  Else
    If MCIa_SendCommand("put " & DeviceAlias & " source at " & _
                       CStr(Source_Left) & " " & _
                       CStr(Source_Top) & " " & _
                       CStr(Source_Width) & " " & _
                       CStr(Source_Height), _
                       , , hCallbackProc, WaitToComplete, ShowErrors) = False Then Exit Function
  End If
  
  ' Move and resize the window to the supplied dimentions
  If MoveWindowLocation = True Then
    If MCIa_SendCommand("put " & DeviceAlias & " window at " & _
                       CStr(Window_Left) & " " & _
                       CStr(Window_Top) & " " & _
                       CStr(Window_Width) & " " & _
                       CStr(Window_Height), _
                       , , hCallbackProc, WaitToComplete, ShowErrors) = False Then Exit Function
  End If
  
  ' Set the destination of the video
  If DisplayVideoOnEntireWindow = True Then
    If MCIa_SendCommand("put " & DeviceAlias & " destination", , , hCallbackProc, WaitToComplete, ShowErrors) = False Then Exit Function
  Else
    If MCIa_SendCommand("put " & DeviceAlias & " destination at " & _
                       CStr(Dest_Left) & " " & _
                       CStr(Dest_Top) & " " & _
                       CStr(Dest_Width) & " " & _
                       CStr(Dest_Height), _
                       , , hCallbackProc, WaitToComplete, ShowErrors) = False Then Exit Function
  End If
  
  ' Operations completed successfully
  MCIa_MoveVideo = True
  
End Function

'=============================================================================================================
' MCIa_RecordCancel
'
' Purpose :
' If the "MCIa_RecordWave" method was called, calling this method will cancel the recording set into motion by
' that method.
'
' Param                Use
' ------------------------------------
' (none)
'
' Return
' ------
' (none)
'
'=============================================================================================================
Public Sub MCIa_RecordCancel()
  
  bUserCanceled = True
  
End Sub

'=============================================================================================================
' MCIa_RecordWave
'
' Purpose :
' This function takes the specified input device and records it to a WAVE output and then saves that output
' to the specified .WAV file.
'
' *NOTE - The device specified in the "IN_DeviceAlias" parameter must have already be opened BEFORE being
'         passed to this function.  However, the device specified in the "OUT_DeviceAlias" parameter does
'         not have to be opened before being passed to this function... this function opens it correctly.
'
' *IMPORTANT - Depending on the multimedia drivers you have installed on your computer, the microphone may
'              be required to do the recording to the WAVE file.
'
' Param                   Use
' ------------------------------------
' OUT_DeviceAlias         Specifies the name of the output device to use to record the WAVE file out with.
' OUT_FilePath            Specifies the path of the file to save the output to.
' OUT_BitsPerSample       Specifies the bits per sample to use in the recording process.  This helps determine
'                         the quality of the output.
' OUT_Channels            Specifies the channels to use in the recording process.  This helps determine the
'                         quality of the output.
' OUT_SamplesPerSec       Specifies the samples per second to use in the recording process.  This helps
'                         determine the quality of the output.
' IN_DeviceAlias          Optional. Specifies the source of the recording.  This device must have already
'                         been opened before passing it to this function.
' IN_SpecifyStartPosition Optional. If set to TRUE, the "IN_StartPosition" parameter is used to specify where
'                         in the input media to start recording from.
' IN_StartPosition        Optional. Specifies where in the input media to start recording from.
' IN_SpecifyEndPosition   Optional. If set to TRUE, the "IN_EndPosition" parameter is used to specify where
'                         to stop recording from the input media.
' IN_EndPosition          Optional. Specifies where to stop recording from the input media.
' RecordInterval          Optional. If specified, determines for how long (in seconds) to record for.
'                         Typically the buffer size is set to 4 seconds. With the MCIWAVE device, the minimum
'                         size is 2 seconds and the maximum size is 9 seconds.
' hCallbackProc           Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                         the device completes an action.  In order for this to work, you must specify the
'                         handle to the window in your initial call to MCI_Open, then specify the handle of
'                         that window in subsequent calls to methods like this.
' ShowErrors              Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                         the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCIa_RecordWave(ByVal OUT_DeviceAlias As String, ByVal OUT_FilePath As String, ByVal OUT_BitsPerSample As BitsPerSample, ByVal OUT_Channels As Channels, ByVal OUT_SamplesPerSec As SamplesPerSec, ByVal IN_DeviceAlias As String, Optional ByVal IN_SpecifyStartPosition As Boolean, Optional ByVal IN_StartPosition As Long, Optional ByVal IN_SpecifyEndPosition As Boolean, Optional ByVal IN_EndPosition As Long, Optional ByVal RecordInterval As Long, Optional ByVal hCallbackProc As Long, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  Dim strCurrentString As String
  Dim intBlockAlign    As Integer
  Dim intBytesPerSeec  As Long
  Dim ReturnValue      As Long
  Dim CurStatus        As String
  Dim bFinished        As Boolean
  Dim FileNum          As Integer
  Dim StopPoint        As Long
  
  ' Make sure the output file path is valid
  OUT_FilePath = Trim(OUT_FilePath)
  If OUT_FilePath = "" Then Exit Function
  
  ' Create a temporary file so that the 16bit file name can be found
  ' (doesn't work on files that don't exist)
  On Error Resume Next
  Kill OUT_FilePath
  On Error GoTo 0
  FileNum = FreeFile
  Open OUT_FilePath For Output As FileNum
    Print #FileNum, "0"
  Close
  OUT_FilePath = Get16bitFilePath(OUT_FilePath)
  If OUT_FilePath = "" Then Exit Function
  Kill OUT_FilePath
  
  ' Make sure the device names are valid
  IN_DeviceAlias = Trim(IN_DeviceAlias)
  OUT_DeviceAlias = Trim(OUT_DeviceAlias)
  If IN_DeviceAlias = "" Or OUT_DeviceAlias = "" Then Exit Function
  If UCase(IN_DeviceAlias) = UCase(OUT_DeviceAlias) Then Exit Function
  
  ' Make sure the INPUT and OUTPUT devices are stopped, and that the OUTPUT device is closed
  mciSendString "stop " & IN_DeviceAlias & Chr(0), vbNullString, 0, 0
  mciSendString "stop " & OUT_DeviceAlias & Chr(0), vbNullString, 0, 0
  mciSendString "close " & OUT_DeviceAlias & Chr(0), vbNullString, 0, 0
  
  ' Open the OUTPUT device to record from
  If RecordInterval > 0 Then
    strCurrentString = "open new type waveaudio alias " & OUT_DeviceAlias & " buffer " & CStr(RecordInterval)
  Else
    strCurrentString = "open new type waveaudio alias " & OUT_DeviceAlias
  End If
  GoSub SendTheString
  
  ' Do some calculations to get the right settings
  intBlockAlign = OUT_Channels * OUT_BitsPerSample / 8
  intBytesPerSeec = OUT_SamplesPerSec * intBlockAlign
  
  ' Setup the OUTPUT device to the specified quality
  strCurrentString = "set " & OUT_DeviceAlias & " " & _
                     "samplespersec " & CStr(OUT_SamplesPerSec) & " " & _
                     "channels " & CStr(OUT_Channels) & " " & _
                     "bitspersample " & CStr(OUT_BitsPerSample) & " " & _
                     "alignment " & CStr(intBlockAlign) & " " & _
                     "bytespersec " & CStr(intBytesPerSeec)
  GoSub SendTheString
  
  ' Set the time format for the INPUT and OUTPUT device
  strCurrentString = "set " & IN_DeviceAlias & " time format milliseconds"
  GoSub SendTheString
  strCurrentString = "set " & OUT_DeviceAlias & " time format milliseconds"
  GoSub SendTheString
  
  ' Get the length of the file to record
  strCurrentString = "seek " & IN_DeviceAlias & " to end"
  GoSub SendTheString
  CurStatus = String(260, Chr(0))
  If mciSendString("status " & IN_DeviceAlias & " position" & Chr(0), CurStatus, 260, 0) <> 0 Then GoTo CancelOut
  CurStatus = Left(CurStatus, InStr(CurStatus, Chr(0)) - 1)
  If CurStatus <> "" Then
    StopPoint = CLng(CurStatus)
    IN_EndPosition = Abs(IN_EndPosition)
    If IN_SpecifyEndPosition = True And IN_EndPosition <> 0 Then
      If IN_EndPosition > (StopPoint - 1) Then IN_EndPosition = StopPoint - 1
    Else
      IN_EndPosition = StopPoint - 1
    End If
  End If
  
  ' If the user specified to start at a certain spot, set the INPUT device to it
  If IN_SpecifyStartPosition = True Then
    strCurrentString = "seek " & IN_DeviceAlias & " to " & CStr(IN_StartPosition)
  Else
    strCurrentString = "seek " & IN_DeviceAlias & " to start"
  End If
  GoSub SendTheString
  
  ' Set the OUTPUT device to start recording
  strCurrentString = "record " & OUT_DeviceAlias & " overwrite"
  GoSub SendTheString
  
  ' Play the INPUT device to start recording
  strCurrentString = "play " & IN_DeviceAlias
  If IN_SpecifyEndPosition = True Then
    strCurrentString = strCurrentString & " to " & CStr(IN_EndPosition)
  End If
  GoSub SendTheString
  
  ' Go into a loop until the play is finished or until the user cancels
  bUserCanceled = False
  bFinished = False
  Do While bUserCanceled = False And bFinished = False
    If bUserCanceled = True Then GoTo CancelOut
    CurStatus = String(260, Chr(0))
    If mciSendString("status " & IN_DeviceAlias & " position" & Chr(0), CurStatus, 260, 0) <> 0 Then GoTo CancelOut
    CurStatus = Left(CurStatus, InStr(CurStatus, Chr(0)) - 1)
    If CLng(CurStatus) >= IN_EndPosition Then bFinished = True
    DoEvents
  Loop
  
  ' Now that the recording is finished, stop the INPUT and OUTPUT devices and save the
  ' recorded WAVE file out to the specified location
  mciSendString "stop " & IN_DeviceAlias & Chr(0), vbNullString, 0, 0
  mciSendString "seek " & IN_DeviceAlias & "to start" & Chr(0), vbNullString, 0, 0
  mciSendString "stop " & OUT_DeviceAlias & Chr(0), vbNullString, 0, 0
  mciSendString "save " & OUT_DeviceAlias & " " & OUT_FilePath & Chr(0), vbNullString, 0, 0
  mciSendString "close " & OUT_DeviceAlias & Chr(0), vbNullString & Chr(0), 0, 0
  
  MCIa_RecordWave = True
  
  Exit Function
  
CancelOut:
  
  mciSendString "stop " & IN_DeviceAlias & Chr(0), vbNullString, 0, 0
  mciSendString "seek " & IN_DeviceAlias & "to start" & Chr(0), vbNullString, 0, 0
  mciSendString "stop " & OUT_DeviceAlias & Chr(0), vbNullString, 0, 0
  mciSendString "close " & OUT_DeviceAlias & Chr(0), vbNullString, 0, 0
  Exit Function
  
SendTheString:
  
  ReturnValue = mciSendString(strCurrentString & Chr(0), vbNullString, 0, hCallbackProc)
  If ReturnValue <> 0 Then
    GetLastErr_MCI ReturnValue, "mciSendString", , , ShowErrors
    GoTo CancelOut
  End If
  Return
  
End Function

'=============================================================================================================
' MCIa_SendCommand
'
' Purpose :
' This function allows you to send commands right to the MCI multimedia system.
'
' Param                Use
' ------------------------------------
' CommandString        Specifies the command to send to the MCI system.
' UseReturnString      Optional. If set to TRUE, this function attempts to caption a return value.  Not all
'                      commands have a return value, but if one is returned, it will be returned via
'                      the "Return_Result" parameter.
' Return_Result        Optional. If the "UseReturnString" parameter is set to TRUE, then this parameter returns
'                      whatever the MCI system returned from the API call.
' hCallbackProc        Optional. Specifies the handle of the window to recieve a MM_MCINOTIFY message when
'                      the device completes an action.  In order for this to work, you must specify the
'                      handle to the window in your initial call to MCI_Open, then specify the handle of
'                      that window in subsequent calls to methods like this.
' WaitToComplete       Optional. If this is set to TRUE, this function will not return control to the application
'                      until the MCI call has completed.
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function MCIa_SendCommand(ByVal CommandString As String, Optional ByVal UseReturnString As Boolean = False, Optional ByRef Return_Result As String, Optional ByVal hCallbackProc As Long, Optional ByVal WaitToComplete As Boolean = True, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  Dim ReturnValue As Long
  
  ' Set the default return value
  Return_Result = ""
  
  ' Make sure parameters are valid
  CommandString = Trim(CommandString)
  If CommandString = "" Then Exit Function
  If InStr(CommandString, Chr(0)) > 0 Then CommandString = Left(CommandString, InStr(CommandString, Chr(0)) - 1)
  If hCallbackProc <> 0 And InStr(UCase(CommandString), " NOTIFY") = 0 Then CommandString = CommandString & " notify"
  If WaitToComplete = True And InStr(UCase(CommandString), " WAIT") = 0 Then CommandString = CommandString & " wait"
  CommandString = CommandString & Chr(0)
  
  ' Call the mciSendString API with the specified command string
  If UseReturnString = False Then
    ReturnValue = mciSendString(CommandString, 0, 0, hCallbackProc)
  Else
    Return_Result = String(260, Chr(0))
    ReturnValue = mciSendString(CommandString, Return_Result, 260, hCallbackProc)
    Return_Result = Left(Return_Result, InStr(Return_Result, Chr(0)) - 1)
  End If
  
  ' Check for errors
  If GetLastErr_MCI(ReturnValue, "mciSendString", , , ShowErrors) = False Then MCIa_SendCommand = True
  
End Function

' Gets or sets the playback volume of any CD audio being played
Public Property Get MCIa_VolumeCD() As Long
  
  Volume_Get hMixer, VolCtrl_CD, MCIa_VolumeCD
  
End Property
Public Property Let MCIa_VolumeCD(ByVal NewValue As Long)
  
  ' Make sure value passed is valid
  If NewValue > VolMax_CD Then
    NewValue = VolMax_CD
  ElseIf NewValue < VolMin_CD Then
    NewValue = VolMin_CD
  End If
  
  ' Set the mixer's volume
  Volume_Set hMixer, VolCtrl_CD, NewValue
  
End Property

' Returns the maximum value that is able to be set via the MCIa_VolumeCD property
Public Property Get MCIa_VolumeCD_Max() As Long
  
  MCIa_VolumeCD_Max = VolMax_CD
  
End Property

' Returns the minimum value that is able to be set via the MCIa_VolumeCD property
Public Property Get MCIa_VolumeCD_Min() As Long
  
  MCIa_VolumeCD_Min = VolMin_CD
  
End Property

' Gets/Sets the playback volume of any MIDI audio being played
Public Property Get MCIa_VolumeMIDI() As Long
  
  Volume_Get hMixer, VolCtrl_Mid, MCIa_VolumeMIDI
  
End Property
Public Property Let MCIa_VolumeMIDI(ByVal NewValue As Long)
  
  ' Make sure value passed is valid
  If NewValue > VolMax_Mid Then
    NewValue = VolMax_Sys
  ElseIf NewValue < VolMin_Mid Then
    NewValue = VolMin_Sys
  End If
  
  ' Set the mixer's volume
  Volume_Set hMixer, VolCtrl_Mid, NewValue
  
End Property

' Returns the maximum value that is able to be set via the MCIa_VolumeMIDI property
Public Property Get MCIa_VolumeMIDI_Max() As Long
  
  MCIa_VolumeMIDI_Max = VolMax_Mid
  
End Property

' Returns the minimum value that is able to be set via the MCIa_VolumeMIDI property
Public Property Get MCIa_VolumeMIDI_Min() As Long
  
  MCIa_VolumeMIDI_Min = VolMin_Mid
  
End Property

' Gets/Sets or sets the master volume of the system.  This effects CD audio, WAVE, MIDI, and all other play back types.
Public Property Get MCIa_VolumeSystem() As Long
  
  Volume_Get hMixer, VolCtrl_Sys, MCIa_VolumeSystem
  
End Property
Public Property Let MCIa_VolumeSystem(ByVal NewValue As Long)
  
  ' Make sure value passed is valid
  If NewValue > VolMax_Sys Then
    NewValue = VolMax_Sys
  ElseIf NewValue < VolMin_Sys Then
    NewValue = VolMin_Sys
  End If
  
  ' Set the mixer's volume
  Volume_Set hMixer, VolCtrl_Sys, NewValue
  
End Property

' Returns the maximum value that is able to be set via the MCIa_VolumeSystem property
Public Property Get MCIa_VolumeSystem_Max() As Long
  
  MCIa_VolumeSystem_Max = VolMax_Sys
  
End Property

' Returns the minimum value that is able to be set via the MCIa_VolumeSystem property
Public Property Get MCIa_VolumeSystem_Min() As Long
  
  MCIa_VolumeSystem_Min = VolMin_Sys
  
End Property

' Gets/Sets or sets the playback volume of any WAVE audio being played
Public Property Get MCIa_VolumeWAVE() As Long
  
  Volume_Get hMixer, VolCtrl_Wav, MCIa_VolumeWAVE
  
End Property
Public Property Let MCIa_VolumeWAVE(ByVal NewValue As Long)
  
  ' Make sure value passed is valid
  If NewValue > VolMax_Wav Then
    NewValue = VolMax_Sys
  ElseIf NewValue < VolMin_Wav Then
    NewValue = VolMin_Sys
  End If
  
  ' Set the mixer's volume
  Volume_Set hMixer, VolCtrl_Wav, NewValue
  
End Property

' Returns the maximum value that is able to be set via the MCIa_VolumeWAVE property
Public Property Get MCIa_VolumeWAVE_Max() As Long
  
  MCIa_VolumeWAVE_Max = VolMax_Wav
  
End Property

' Returns the minimum value that is able to be set via the MCIa_VolumeWAVE property
Public Property Get MCIa_VolumeWAVE_Min() As Long
  
  MCIa_VolumeWAVE_Min = VolMin_Wav
  
End Property



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



'=============================================================================================================
' RES_Extract
'
' Purpose :
' This function extracts the specified resource from the current project's included resource file.
'
' Param                Use
' ------------------------------------
' Return_Resource      Returns the specified resource in the appropriate form.  Picture for Bitmap, Icon, and
'                      cursor resources, string for string resources, and byte array for data (binary) resources.
' RES_ID               Specifies the ID of the resource to return.
' RES_Type             Specifies what type of data the resource is.
' RES_Name             Optional. Specifies the name that the resource is stored under.  This only applies for
'                      "Data" resources.  The default is "CUSTOM".
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function RES_Extract(ByRef Return_Resource As Variant, ByVal RES_ID As Long, ByVal RES_Type As ResDataTypes, Optional ByVal RES_Name As String = "CUSTOM", Optional ByVal ShowErrors As Boolean = True) As Boolean
On Error GoTo ErrorTrap
  
  ' Reset the return variable(s)
  Return_Resource = Empty
  
  ' Check if a Type name has been specified
  RES_Name = Trim(RES_Name)
  If RES_ID < 1 Or RES_ID > 32767 Then
    If ShowErrors = True Then MsgBox "Invalid resource ID specified.  Can't retrieve the file from the resource.", vbOKOnly + vbExclamation, "  Invalid Resource ID Specified"
    Exit Function
  End If
  
  ' Load the resource
  Select Case RES_Type
    
    ' BITMAP, CURSOR, ICON
    Case rdt_Bitmap, rdt_Cursor, rdt_Icon
      
      ' Returns an "StdPicture" object.  To clean it up, set it equal to NOTHING (Set Return_Resource = Nothing)
      Set Return_Resource = LoadResPicture(RES_ID, CInt(RES_Type))
      
    ' DATA
    Case rdt_Data
      
      ' Returns a byte array.  To clean it up, erase it (Erase Return_Resource)
      Return_Resource = LoadResData(RES_ID, RES_Name)
      
    ' STRING
    Case rdt_String
      
      ' Returns a string.  To clean it up, set it equal to a blank string (Return_Resource = "")
      Return_Resource = LoadResString(RES_ID)
      
  End Select
  
  RES_Extract = True
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    If ShowErrors = True Then MsgBox "The following error occured while trying to extract a file from resource:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
  End If
  
End Function

'=============================================================================================================
' RES_Save
'
' Purpose :
' This function extracts the specified resource from the current project's included resource file and saves
' it to the specified file path.
'
' Param                Use
' ------------------------------------
' FileToExtractTo      Specifies the path of the file to save the resource to.  The function properly saves
'                      the resource based on it's type.  Bitmap file (.BMP) for bitmap resources, Icon file
'                      (.ICO) for icon resources, Cursor file (.CUR) for cursor resources, ASCII Text files
'                      (.TXT) for string resources, and binary files for Data resources.
' RES_ID               Specifies the ID of the resource to return.
' RES_Type             Specifies what type of data the resource is.
' RES_Name             Optional. Specifies the name that the resource is stored under.  This only applies for
'                      "Data" resources.  The default is "CUSTOM".
' ShowErrors           Optional. If this is set to TRUE and an error occurs, the user will be shown a message describing
'                      the error that occured.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function RES_Save(ByVal FileToExtractTo As String, ByVal RES_ID As Long, ByVal RES_Type As ResDataTypes, Optional ByVal RES_Name As String = "CUSTOM", Optional ByVal OverwriteIfExists As Boolean = True, Optional ByVal ShowErrors As Boolean = True) As Boolean
On Error GoTo ErrorTrap
  
  Dim RESPic    As StdPicture
  Dim ResFile() As Byte
  Dim RESString As String
  Dim FileNum   As Long
  
  ' Check if a Type name has been specified
  FileToExtractTo = Trim(FileToExtractTo)
  RES_Name = Trim(RES_Name)
  If FileToExtractTo = "" Then
    If ShowErrors = True Then MsgBox "No file specified to extract resource to.", vbOKOnly + vbExclamation, "  No Output File Specified"
    Exit Function
  ElseIf RES_ID < 1 Or RES_ID > 32767 Then
    If ShowErrors = True Then MsgBox "Invalid resource ID specified.  Can't retrieve the file from the resource.", vbOKOnly + vbExclamation, "  Invalid Resource ID Specified"
    Exit Function
  End If
  
  ' Check if the file exists, and if it does, overwrite it
  If FileExists(FileToExtractTo) = True Then
    If OverwriteIfExists = True Then
      Kill FileToExtractTo
    Else
      RES_Save = True
      Exit Function
    End If
  End If
  
  ' Load the resource
  Select Case RES_Type
    
    ' BITMAP, CURSOR, ICON
    Case rdt_Bitmap, rdt_Cursor, rdt_Icon
      Set RESPic = LoadResPicture(RES_ID, CInt(RES_Type))
      If Not RESPic Is Nothing Then
        SavePicture RESPic, FileToExtractTo
        Set RESPic = Nothing
      End If
      
    ' DATA
    Case rdt_Data
      ResFile = LoadResData(RES_ID, RES_Name)
      If UBound(ResFile) > 0 Then
        FileNum = FreeFile
        Open FileToExtractTo For Binary As #FileNum
          Put #FileNum, 1, ResFile()
        Close #FileNum
        Erase ResFile
      End If
      
    ' STRING
    Case rdt_String
      RESString = LoadResString(RES_ID)
      If RESString <> "" Then
        FileNum = FreeFile
        Open FileToExtractTo For Output As #FileNum
          Print #FileNum, RESString
        Close #FileNum
        RESString = ""
      End If
      
  End Select
  
  RES_Save = True
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    If ShowErrors = True Then MsgBox "The following error occured while trying to save a file from resource:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
  End If
  
End Function



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



'=============================================================================================================
' OS_GetCdCount
'
' Purpose :
' This function goes through all mapped and local drives and returns back the number of CD-ROM drives found.
'
' Param                Use
' ------------------------------------
' (none)
'
' Return
' ------
' The number of CD-ROM devices found
'
'=============================================================================================================
Public Function OS_GetCdCount() As Integer
On Error Resume Next
  
  Dim MyCounter As Integer
  Dim CDDrives As Integer
  
  ' Chr(65) = "A"   Chr(90) = "Z"
  For MyCounter = 65 To 90
    If OS_CheckDriveType(Chr(MyCounter)) = "CD-ROM" Then
      CDDrives = CDDrives + 1
    End If
  Next
  
  OS_GetCdCount = CDDrives
  
End Function

'=============================================================================================================
' OS_CheckDriveType
'
' Purpose :
' This function takes the specified drive leter and checks what type of drive it is.
'
' Param                Use
' ------------------------------------
' DriveLetter          Specifies the letter (A - Z) of the local or mapped drive to check.
'
' Return
' ------
' A string representing the type of drive found at the specified letter.
'
'=============================================================================================================
Public Function OS_CheckDriveType(ByVal DriveLetter As String) As String
On Error Resume Next
  
  Dim lngReturnValue As Long
  
  If Len(DriveLetter) > 1 Or DriveLetter = "" Then OS_CheckDriveType = "Invalid Drive Letter"
  
  lngReturnValue = GetDriveTypeA(UCase(DriveLetter) & ":\")

  Select Case lngReturnValue
    Case 0 ' Cannot be determined
      OS_CheckDriveType = "Unknown"
    Case 1 ' The root directory does not exist
      OS_CheckDriveType = "Unknown"
    Case DRIVE_CDROM
      OS_CheckDriveType = "CD-ROM"
    Case DRIVE_REMOVABLE
      OS_CheckDriveType = "Removable"
    Case DRIVE_FIXED
      OS_CheckDriveType = "Fixed"
    Case DRIVE_REMOTE
      OS_CheckDriveType = "Remote"
  End Select
  
End Function



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



' This function attempts to obtain a mixer control for the specified component
Private Function GetVolumeObject(ByVal hMixer As Long, ByVal ComponentType As MixerComponentTypes, ByRef Return_MIXERCONTROL As MIXERCONTROL) As Boolean
On Error Resume Next
  
  Dim ReturnCode As Long
  Dim MxLC       As MIXERLINECONTROLS
  Dim MxL        As MIXERLINE
  Dim hMem       As Long
  
  MxL.cbStruct = Len(MxL)
  MxL.dwComponentType = ComponentType
  
  ' Get a line corresponding to the component type
  ReturnCode = mixerGetLineInfo(hMixer, MxL, MIXER_GETLINEINFOF_COMPONENTTYPE)
  If ReturnCode = MMSYSERR_NOERROR Then
    With MxLC
      .cbStruct = Len(MxLC)
      .dwLineID = MxL.dwLineID
      .dwControl = MIXERCONTROL_CONTROLTYPE_VOLUME
      .cControls = 1
      .cbmxctrl = Len(Return_MIXERCONTROL)
    End With
    
    ' Allocate a buffer for the control
    hMem = GlobalAlloc(&H40, Len(Return_MIXERCONTROL))
    MxLC.pamxctrl = GlobalLock(hMem)
    Return_MIXERCONTROL.cbStruct = Len(Return_MIXERCONTROL)
    
    ' Get the control
    ReturnCode = mixerGetLineControls(hMixer, MxLC, MIXER_GETLINECONTROLSF_ONEBYTYPE)
    If ReturnCode = MMSYSERR_NOERROR Then
      GetVolumeObject = True
      CopyStructFromPtr Return_MIXERCONTROL, MxLC.pamxctrl, Len(Return_MIXERCONTROL) ' Copy the control into the destination structure
    End If
    
    GlobalFree hMem
  End If
  
End Function

' Gets the volume of the specified component
Private Function Volume_Get(ByVal hMixer As Long, ByRef ControlToUse As MIXERCONTROL, ByRef Return_Volume As Long) As Boolean
On Error Resume Next
  
  Dim MxCD As MIXERCONTROLDETAILS
  
  Return_Volume = 0
  With MxCD
    .cbStruct = Len(MxCD)
    .dwControlID = ControlToUse.dwControlID
    .cChannels = 1
    .cbDetails = Len(Return_Volume)
    .paDetails = VarPtr(Return_Volume)
  End With
  
  ' Get the control's volume
  If mixerGetControlDetails(hMixer, MxCD, MIXER_GETCONTROLDETAILSF_VALUE) = MMSYSERR_NOERROR Then Volume_Get = True
  
End Function

' Sets the volume of the specified component
Private Function Volume_Set(ByVal hMixer As Long, ByRef ControlToUse As MIXERCONTROL, ByVal lngVolume As Long) As Boolean
On Error Resume Next
  
  Dim MxCD As MIXERCONTROLDETAILS
  
  With MxCD
    .cbStruct = Len(MxCD)
    .dwControlID = ControlToUse.dwControlID
    .cChannels = 1
    .cbDetails = Len(lngVolume)
    .paDetails = VarPtr(lngVolume)
  End With
  
  ' Set the control's volume
  If mixerSetControlDetails(hMixer, MxCD, MIXER_SETCONTROLDETAILSF_VALUE) = MMSYSERR_NOERROR Then Volume_Set = True
  
End Function

' Function that takes a 32bit file path and converts it to the 16bit equivelant file path
Private Function Get16bitFilePath(ByVal str32bitFilePath As String) As String
On Error Resume Next
  
  str32bitFilePath = Trim(str32bitFilePath)
  If str32bitFilePath = "" Then Exit Function
  If Right(str32bitFilePath, 1) <> Chr(0) Then str32bitFilePath = str32bitFilePath & Chr(0)
  
  Get16bitFilePath = String(260, Chr(0))
  If GetShortPathName(str32bitFilePath, Get16bitFilePath, 260) = 0 Then
    Get16bitFilePath = ""
  Else
    Get16bitFilePath = UCase(Left(Get16bitFilePath, InStr(Get16bitFilePath, Chr(0)) - 1))
  End If
  
End Function

' Get the error that just occured (if one did occur)
Private Function GetLastErr_MCI(Optional ByVal ErrorNumber As Long, Optional ByVal LastAPICalled As String = "last", Optional ByRef Return_ErrNum As Long, Optional ByRef Return_ErrDesc As String, Optional ByVal ShowErrors As Boolean = False) As Boolean
  
  Dim intErrorNumber As Integer
  
  ' Clear return variables
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' If no error occured, exit
  If ErrorNumber = 0 Then
    Exit Function
  Else
    intErrorNumber = LOWORD(ErrorNumber)
  End If
  
  ' Assign initial values
  GetLastErr_MCI = True
  Return_ErrNum = CLng(intErrorNumber)
  If LastAPICalled = "" Then LastAPICalled = "last"
  
  ' Get the error description based on the error number
  Return_ErrDesc = String(260, Chr(0))
  If mciGetErrorString(Return_ErrNum, Return_ErrDesc, 260) <> 0 Then

    Return_ErrDesc = Left(Return_ErrDesc, InStr(Return_ErrDesc, Chr(0)) - 1)
  Else
    Return_ErrDesc = "Unknown Error"
  End If
  
  ' If specified to show an error message, do so
  If ShowErrors = True Then
    MsgBox "The following error occured while calling the " & LastAPICalled & " MCI API:" & Chr(13) & Chr(13) & _
           "Error Number = " & Return_ErrNum & Chr(13) & _
           "Error Description = " & Return_ErrDesc, vbOKOnly + vbExclamation, "  MCI ERROR"
  End If
  
End Function

' Function that extracts the "Low Order" (lower 16bits) of a 32bit number
Private Function LOWORD(ByVal dwValue As Long) As Integer
  
  LOWORD = Val("&H" & Right("0000" & Hex(dwValue), 4))
  
End Function

' Function that extracts the "High Order" (upper 16bits) of a 32bit number
Private Function HIWORD(ByVal dwValue As Long) As Integer
  
  HIWORD = Val("&H" & Left(Right("00000000" & Hex(dwValue), 8), 4))
  
End Function

' Function that checks if a file exists
Private Function FileExists(ByVal FilePath As String) As Boolean
  
  If Dir(FilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
    FileExists = True
  Else
    FileExists = False
  End If
  
End Function
