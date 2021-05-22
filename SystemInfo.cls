VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' System Information class

' Version API Structure
Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersion As Long
    dwFileVersionMS As Long
    dwFileVersionLS As Long
    dwProductVersionMS As Long
    dwProductVersionLS As Long
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize  As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private osvi As OSVERSIONINFOEX

Public Enum VersionTypeMask
    VER_MINORVERSION = &H1
    VER_MAJORVERSION = &H2
    VER_BUILDNUMBER = &H4
    VER_PLATFORMID = &H8
    VER_SERVICEPACKMINOR = &H10
    VER_SERVICEPACKMAJOR = &H20
    VER_SUITENAME = &H40
    VER_PRODUCT_TYPE = &H80
End Enum

Public Enum ComparionTypes
    VER_EQUAL = 1
    VER_GREATER = 2
    VER_GREATER_EQUAL = 3
    VER_LESS = 4
    VER_LESS_EQUAL = 5
    VER_AND = 6
    VER_OR = 7
End Enum

Private Declare Function GetVersionEx _
 Lib "kernel32" Alias "GetVersionExA" _
 (lpVersionInformation As Any) As Long
 
Private Declare Function GetFileVersionInfoSize _
 Lib "version.dll" Alias "GetFileVersionInfoSizeA" _
 (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
 
Private Declare Sub CopyMemory _
 Lib "kernel32" Alias "RtlMoveMemory" _
 (Destination As Any, Source As Any, ByVal Length As Long)
 
Private Declare Function GetFileVersionInfo _
 Lib "version.dll" Alias "GetFileVersionInfoA" _
 (ByVal lptstrFilename As String, ByVal dwHandle As Long, _
 ByVal dwLen As Long, lpData As Any) As Long
 
Private Declare Function VerQueryValue _
 Lib "version.dll" Alias "VerQueryValueA" _
(pBlock As Any, ByVal lpSubBlock As String, _
lplpBuffer As Long, puLen As Long) As Long

Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const VER_NT_WORKSTATION = &H1
Private Const VER_NT_DOMAIN_CONTROLLER = &H2
Private Const VER_NT_SERVER = &H3

Private Declare Function RegCloseKey _
 Lib "advapi32.dll" (ByVal hKey As Long) As Long
 
Private Declare Function RegQueryValueEx _
 Lib "advapi32.dll" Alias "RegQueryValueExA" _
 (ByVal hKey As Long, ByVal lpValueName As String, _
 ByVal lpReserved As Long, lpType As Long, _
 lpData As Any, lpcbData As Long) As Long
 
Private Declare Function RegOpenKeyEx _
 Lib "advapi32.dll" Alias "RegOpenKeyExA" _
 (ByVal hKey As Long, ByVal lpSubKey As String, _
 ByVal ulOptions As Long, ByVal samDesired As Long, _
 phkResult As Long) As Long
    
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_QUERY_VALUE = &H1
' Version APIs end

' CP Info
Private Const ERROR_INVALID_PARAMETER = 87
Private Const MAX_LEADBYTES = 12
Private Const MAX_DEFAULTCHAR = 2
Private Const CP_ACP = 0

Private Type CPINFO
    MaxCharSize As Long                         '  max length (Byte) of a char
    DefaultChar(MAX_DEFAULTCHAR - 1) As Byte    '  default character
    LeadByte(MAX_LEADBYTES - 1) As Byte         '  lead byte ranges
End Type

Private Declare Function GetCPInfo Lib "kernel32" _
 (ByVal CodePage As Long, lpCPInfo As CPINFO) As Long
' CP Info

' SysInfo
Public Enum ProcessorType
    PROCESSOR_ARCHITECTURE_INTEL = 0
    PROCESSOR_ARCHITECTURE_MIPS = 1
    PROCESSOR_ARCHITECTURE_ALPHA = 2
    PROCESSOR_ARCHITECTURE_PPC = 3
    PROCESSOR_ARCHITECTURE_UNKNOWN = &HFFFF
End Enum

Public Enum ExtendedNameFormat
    enfNameUnknown = 0
    enfNameFullyQualifiedDN = 1
    enfNameSamCompatible = 2
    enfNameDisplay = 3
    enfNameUniqueId = 6
    enfNameCanonical = 7
    enfNameUserPrincipal = 8
    enfNameCanonicalEx = 9
    enfNameServicePrincipal = 10
End Enum

Public Enum ComputerNameFormat
    cnfComputerNameNetBIOS = 0
    cnfComputerNameDnsHostname = 1
    cnfComputerNameDnsDomain = 2
    cnfComputerNameDnsFullyQualified = 3
    cnfComputerNamePhysicalNetbios = 4
    cnfComputerNamePhysicalDnsHostname = 5
    cnfComputerNamePhysicalDnsDomain = 6
    cnfComputerNamePhysicalDnsFullyQualified = 7
    cnfComputerNameMax = 8
End Enum

Public Enum ProductSuiteType
    VER_SUITE_BACKOFFICE = &H4      'Microsoft® BackOffice® components are installed.
    VER_SUITE_DATACENTER = &H80     'Windows 2000 Datacenter Server is installed.
    VER_SUITE_ENTERPRISE = &H2      'Windows® 2000 Advanced Server is installed.
    VER_SUITE_SMALLBUSINESS = &H1      'Microsoft® Small Business Server is installed.
    VER_SUITE_SMALLBUSINESS_RESTRICTED = &H20     ' Microsoft® Small Business Server is installed with the restrictive client license in force.
    VER_SUITE_TERMINAL = &H10    ' Terminal Services is installed.
End Enum

Private Type SYSTEM_INFO
    'dwOemID As Long            'Obsolete, use Union instead
    wProcessorArchitecture As Integer
    wReserved As Integer      'Reserved
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type
Private si As SYSTEM_INFO

Private Const SM_SECURE = 44
Private Const SM_NETWORK = 63
Private Const SM_CLEANBOOT = 67
Private Const SM_SLOWMACHINE = 73
Private Const SM_MIDEASTENABLED = 74
Private Const SM_IMMENABLED = 82
Private Const SM_REMOTESESSION = &H1000
Private Const SM_SHOWSOUNDS = 70

' SystemParametersInfo flags
Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPIF_SENDWININICHANGE = &H2

' This is a made-up constant.
Private Const SPIF_TELLALL = SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE

Private Const SPI_GETBEEP = 1
Private Const SPI_SETBEEP = 2
Private Const SPI_GETSCREENSAVETIMEOUT = 14
Private Const SPI_SETSCREENSAVETIMEOUT = 15
Private Const SPI_GETSCREENSAVEACTIVE = 16
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const SPI_GETSCREENSAVERRUNNING = 114
Private Const SPI_GETWINDOWSEXTENSION = 92

Private Declare Function SystemParametersInfo _
 Lib "user32" Alias "SystemParametersInfoA" _
 (ByVal uAction As Long, ByVal uParam As Long, _
 lpvParam As Any, ByVal fuWinIni As Long) As Long
 
Private Declare Function GetComputerName _
 Lib "kernel32" Alias "GetComputerNameA" _
 (ByVal lpBuffer As String, nSize As Long) As Long
 
Private Declare Function GetComputerNameEx _
 Lib "kernel32" Alias "GetComputerNameExA" _
 (ByVal NameType As ComputerNameFormat, ByVal lpBuffer As String, _
 nSize As Long) As Long
 
Private Declare Function SetComputerName _
 Lib "kernel32" Alias "SetComputerNameA" _
 (ByVal lpComputerName As String) As Long
 
Private Declare Function SetComputerNameEx _
 Lib "kernel32" _
 (ByVal NameType As ComputerNameFormat, _
 ByVal lpBuffer As String) As Long
 
Private Declare Function GetUserName _
 Lib "advapi32.dll" Alias "GetUserNameA" _
 (ByVal lpBuffer As String, nSize As Long) As Long
 
Private Declare Function GetUserNameEx _
 Lib "secur32.dll" Alias "GetUserNameExA" _
 (ByVal NameFormat As Long, ByVal lpNameBuffer As String, _
 nSize As Long) As Long
 
Private Declare Function GetWindowsDirectory _
 Lib "kernel32" Alias "GetWindowsDirectoryA" _
 (ByVal lpBuffer As String, ByVal nSize As Long) As Long
 
Private Declare Function GetSystemDirectory _
 Lib "kernel32" Alias "GetSystemDirectoryA" _
 (ByVal lpBuffer As String, ByVal nSize As Long) As Long
 
Private Declare Function GetTempPath _
 Lib "kernel32" Alias "GetTempPathA" _
 (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
 
Private Declare Sub GetSystemInfo _
 Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
 
Private Declare Function GetSystemMetrics _
 Lib "user32" (ByVal nIndex As Long) As Long
 
Private Declare Function FormatMessage _
 Lib "kernel32" Alias "FormatMessageA" _
 (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
 ByVal dwLanguageId As Long, ByVal lpBuffer As String, _
 ByVal nSize As Long, Arguments As Long) As Long

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Private Const dhcMaxComputerName = 15
Private Const dhcMaxPath = 260

' SysInfo End

Public Enum siCSIDL_VALUES
    CSIDL_FLAG_CREATE = &H8000 ' (Version 5.0)
    CSIDL_ADMINTOOLS = &H30 ' (Version 5.0)
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_APPDATA = &H1A ' (Version 4.71)
    CSIDL_BITBUCKET = &HA
    CSIDL_COMMON_ADMINTOOLS = &H2F  ' Version 5
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_APPDATA = &H23  ' Version 5
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_CONTROLS = &H3
    CSIDL_COOKIES = &H21
    CSIDL_DESKTOP = &H0
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_FAVORITES = &H6
    CSIDL_FONTS = &H14
    CSIDL_HISTORY = &H22
    CSIDL_INTERNET = &H1
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_LOCAL_APPDATA = &H1C      ' Version 5
    CSIDL_MYPICTURES = &H27  ' Version 5
    CSIDL_NETHOOD = &H13
    CSIDL_NETWORK = &H12
    CSIDL_PERSONAL = &H5
    CSIDL_PRINTERS = &H4
    CSIDL_PRINTHOOD = &H1B
    CSIDL_PROFILE = &H28  ' Version 5
    CSIDL_PROGRAM_FILES = &H2A  ' Version 5
    CSIDL_PROGRAM_FILES_COMMON = &H2B  ' Version 5
    CSIDL_PROGRAMS = &H2
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_STARTMENU = &HB
    CSIDL_STARTUP = &H7
    CSIDL_SYSTEM = &H25  ' Version 5
    CSIDL_TEMPLATES = &H15
    CSIDL_WINDOWS = &H24    ' Version 5.0.
End Enum

Private Declare Function SHGetSpecialFolderLocation _
 Lib "shell32" _
 (ByVal hwndOwner As Long, ByVal nFolder As Long, _
 ppidl As Long) As Long

Private Declare Function SHGetPathFromIDList _
 Lib "shell32" _
 (pidl As Long, ByVal pszPath As String) As Long

Private Declare Sub CoTaskMemFree Lib "ole32" _
 (ByVal pv As Long)
    
Private Const MAX_PATH = 260
Private Const NOERROR = 0


Private mblnVersionInfoEx As Boolean

' 5113 is arbitrary.
Private Const dhcErrBase = vbObjectError + 5113
Private Const ERR_STRING = "Invalid for this operating system."
Private Const ERR_INVALID_OS = dhcErrBase + 1
Private Const ERR_NAME_TOO_LONG = dhcErrBase + 2
Private Const ERR_INVALID_NAME = dhcErrBase + 3

' Should this class raise errors if the
' operating system doesn't support the
' requested operation, or should it silently fail?
Public RaiseErrors As Boolean

Public Property Get BootMethod() As Long
    ' Retrieve the boot method.
    ' 0 = Normal boot
    ' 1 = Fail-safe boot
    ' 2 = Fail-safe with network boot
    BootMethod = GetSystemMetrics(SM_CLEANBOOT)
End Property

Public Property Get MidEastEnabled() As Boolean
    ' Returns True if the system is enabled for
    ' Hebrew/Arabic languages.
    MidEastEnabled = CBool(GetSystemMetrics(SM_MIDEASTENABLED))
End Property

Public Property Get NetworkPresent() As Boolean
    ' Returns True if a network is present.
    ' Check the least-significant bit to see if
    ' a network is installed.
    NetworkPresent = CBool(GetSystemMetrics(SM_NETWORK) And 1)
End Property

Public Property Get IsIMMEnabled() As Boolean
    ' Windows 2000: TRUE if Input Method Manager/Input Method Editor
    ' features are enabled; FALSE otherwise.
    ' SM_IMMENABLED can determine if the system handles Unicode IME.
    ' However, if the IME is language-dependent you should also check that
    ' the target language has been installed. Otherwise some components, like
    ' fonts or registry settings, may not be present.
    If Me.IsWin2000 Then
        IsIMMEnabled = CBool(GetSystemMetrics(SM_IMMENABLED))
    Else
        Call HandleErrors(ERR_INVALID_OS)
    End If
End Property

Public Property Get IsRemoteSession() As Boolean
    ' Windows NT 4.0 SP4 or later: This system metric is used in a
    ' Terminal Services environment. If the calling process is associated with a
    ' Terminal Services client session, the return value is TRUE. If the calling process
    ' is associated with the Terminal Server console session, the return value is zero.
    IsRemoteSession = CBool(GetSystemMetrics(SM_REMOTESESSION))
End Property

Public Property Get Secure() As Boolean
    ' Returns True if security is present.
    Secure = CBool(GetSystemMetrics(SM_SECURE))
End Property

Public Property Get ShowSounds() As Boolean
    ' True if the user requires an application to present information
    ' visually in situations where it would otherwise present the information
    ' only in audible form
    ShowSounds = CBool(GetSystemMetrics(SM_SHOWSOUNDS))
End Property

Public Property Get SlowMachine() As Boolean
    ' Returns True if computer has a low-end processor.
    SlowMachine = CBool(GetSystemMetrics(SM_SLOWMACHINE))
End Property

Public Property Let Beep(Value As Boolean)
    ' Turns the system warning beeper on or off.
    Call SystemParametersInfo(SPI_SETBEEP, Value, 0, SPIF_TELLALL)
End Property

Public Property Get Beep() As Boolean
    ' Turns the system warning beeper on or off.
    Dim fBeep As Boolean
    Call SystemParametersInfo(SPI_GETBEEP, 0, fBeep, 0)
    Beep = fBeep
End Property

Public Property Get ScreenSaverActive() As Boolean
    ' Set or retrieve the state of the screen saver.
    Dim lngValue As Long
    Call SystemParametersInfo(SPI_GETSCREENSAVEACTIVE, 0, lngValue, 0)
    ScreenSaverActive = CBool(lngValue)
End Property

Public Property Let ScreenSaverActive(Value As Boolean)
    ' Set or retrieve the state of the screen saver.
    Call SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, Value, 0, SPIF_TELLALL)
End Property

Public Property Get ScreenSaverTimeout() As Long
    ' Set or retrieve the screen saver time-out value.
    Dim lngValue As Long
    Call SystemParametersInfo(SPI_GETSCREENSAVETIMEOUT, 0, lngValue, 0)
    ScreenSaverTimeout = lngValue
End Property

Public Property Let ScreenSaverTimeout(Value As Long)
    ' Set or retrieve the screen saver time-out value.
    Call SystemParametersInfo(SPI_SETSCREENSAVETIMEOUT, Value, 0, SPIF_TELLALL)
End Property

Public Property Get ScreenSaverRunning() As Boolean
    ' Windows 98, Windows 2000: Determines whether a screen saver is currently
    ' running on the window.
    Dim lngRunning As Long
    If IsWin98 Or IsWin2000 Then
        Call SystemParametersInfo(SPI_GETSCREENSAVERRUNNING, 0, lngRunning, 0)
        ScreenSaverRunning = CBool(lngRunning)
    Else
        Call HandleErrors(ERR_INVALID_OS)
    End If
End Property

Public Property Get WindowsExtension() As Boolean
    ' Win95 only. Indicates whether the Windows extension, Windows Plus!, is installed.
    Dim lngValue As Long
    Call SystemParametersInfo(SPI_GETWINDOWSEXTENSION, 0, lngValue, 0)
    WindowsExtension = CBool(lngValue)
End Property

Public Property Get IsDBCS() As Boolean
    ' Is the operating system working with DBCS characters?
    ' If so, you'll need to alter API calls.
    Dim ocpi As CPINFO
    Call GetCPInfo(CP_ACP, ocpi)
    IsDBCS = (ocpi.MaxCharSize > 1)
End Property

Public Property Get TempPath() As String
    ' Retrieve the Windows temporary path.
    ' Windows 95/98: The GetTempPath function gets the temporary file path as follows:
    ' 1)  The path specified by the TMP environment variable.
    ' 2)  The path specified by the TEMP environment variable, if TMP is not
    '         defined or if TMP specifies a directory that does not exist.
    ' 3)  The current directory, if both TMP and TEMP are not defined or
    '         specify nonexistent directories.
    ' WinNT/Win2000: The GetTempPath function does not verify that the
    '     directory specified by the TMP or TEMP environment
    '     variables exists. The function gets the temporary file path as follows:
    ' 1) The path specified by the TMP environment variable.
    ' 2) The path specified by the TEMP environment variable, if TMP is not defined.
    ' 3) The Windows directory, if both TMP and TEMP are not defined.
    '
    Dim strBuffer As String
    Dim lngLen As Long
    
    strBuffer = Space(dhcMaxPath)
    lngLen = dhcMaxPath
    lngLen = GetTempPath(lngLen, strBuffer)
    ' If the path is longer than dhcMaxPath, then
    ' lngLen contains the correct length. Resize the
    ' buffer and try again.
    If lngLen > dhcMaxPath Then
        strBuffer = Space(lngLen)
        lngLen = GetTempPath(lngLen, strBuffer)
    End If
    TempPath = Left$(strBuffer, lngLen)
End Property

Public Property Get WindowsDirectory() As String
    ' Retrieve the Windows directory.
    Dim strBuffer As String
    Dim lngLen As Long
    
    strBuffer = Space(dhcMaxPath)
    lngLen = dhcMaxPath
    lngLen = GetWindowsDirectory(strBuffer, lngLen)
    ' If the path is longer than dhcMaxPath, then
    ' lngLen contains the correct length. Resize the
    ' buffer and try again.
    If lngLen > dhcMaxPath Then
        strBuffer = Space(lngLen)
        lngLen = GetWindowsDirectory(strBuffer, lngLen)
    End If
    WindowsDirectory = Left$(strBuffer, lngLen)
End Property

Public Property Get SystemDirectory() As String
    ' Retrieve the system directory.
    Dim strBuffer As String
    Dim lngLen As Long
    
    strBuffer = Space(dhcMaxPath)
    lngLen = dhcMaxPath
    
    lngLen = GetSystemDirectory(strBuffer, lngLen)
    ' If the path is longer than dhcMaxPath, then
    ' lngLen contains the correct length. Resize the
    ' buffer and try again.
    If lngLen > dhcMaxPath Then
        strBuffer = Space(lngLen)
        lngLen = GetSystemDirectory(strBuffer, lngLen)
    End If
    SystemDirectory = Left$(strBuffer, lngLen)
End Property

Public Property Get ComputerName( _
 Optional NameFormat As ComputerNameFormat = cnfComputerNameNetBIOS) _
 As String

    ' Set or retrieve the NetBIOS name of the computer.
    Dim strBuffer As String
    Dim lngLen As Long

    If IsWin2000 Then
        If NameFormat <> cnfComputerNameNetBIOS Then
            ' If a particular NameFormat is requested and the
            ' OS is Windows 2000, then use the Extended
            ' version of the API function.

            ' To determine the required buffer size for the
            ' particular value of NameFormat, pass vbNullString
            ' for strBuffer. When the function returns, lngLen will
            ' contain the length of the required buffer.
            Call GetComputerNameEx(NameFormat, vbNullString, lngLen)
            strBuffer = String$(lngLen + 1, vbNullChar)
            If CBool(GetComputerNameEx( _
             NameFormat, strBuffer, lngLen)) Then
                ComputerName = Left$(strBuffer, lngLen)
            End If
        Else
            ' Specified NameFormat is cnfComputerNameNetBios
            ' in which case, use GetComputerName API
            strBuffer = String$(dhcMaxComputerName + 1, vbNullChar)
            lngLen = Len(strBuffer)
            If CBool(GetComputerName(strBuffer, lngLen)) Then
                ' If successful, return the buffer
                ComputerName = Left$(strBuffer, lngLen)
            End If
        End If
    Else
        ' The OS is not Win2000
        ' Only cnfComputerNameNetBios is valid for NameFormat
        If NameFormat = cnfComputerNameNetBIOS Then
            strBuffer = String$(dhcMaxComputerName + 1, vbNullChar)
            lngLen = Len(strBuffer)
            If CBool(GetComputerName(strBuffer, lngLen)) Then
                ' If successful, return the buffer
                ComputerName = Left$(strBuffer, lngLen)
            End If
        Else
            If RaiseErrors Then
                Call HandleErrors(ERR_INVALID_OS)
            End If
        End If
    End If
End Property

Public Property Let ComputerName( _
 Optional NameFormat As ComputerNameFormat = cnfComputerNameNetBIOS, _
 Name As String)

    ' SetComputerName changes the registry, not the current
    ' computer name.
    '
    ' Windows 95, Windows 98: If this string contains one or more
    ' characters that are outside the standard character set,
    ' those characters are coerced into standard characters.
    ' Windows NT: If this string contains one or more characters
    ' that are outside the standard character set, SetComputerName
    ' returns ERROR_INVALID_PARAMETER. It does not coerce the characters
    ' outside the standard set.
    ' The standard character set includes letters, numbers, and the
    ' following symbols: ! @ # $ % ^ & ' ) ( . - _ { } ~ .

    If NameFormat <> cnfComputerNameNetBIOS And IsWin2000 Then
        ' If a particular NameFormat is requested and the OS is
        ' Windows 2000, then use the Extended version of the
        ' API function. Requires administrator privileges on
        ' the local computer.
        '
        ' Name cannot include control characters, leading or
        ' trailing spaces, or any of the following characters:
        ' " / \ [ ] : | < > + = ; , ?

        ' Restrictions on NameTypeFormat value:
        '   cnfComputerNamePhysicalNetbios
        '       Sets the NetBIOS name and the DNS host name _
        '       to the name specified in lpBuffer. The name cannot
        '       exceed MAX_COMPUTERNAME_LENGTH characters,
        '       not including the terminating null character.
        '   cnfComputerNamePhysicalDnsHostname
        '       Sets the NetBIOS and the DNS host name name
        '       to the name specified in lpBuffer. If the name exceeds
        '       MAX_COMPUTERNAME_LENGTH characters, the NetBIOS name is
        '       truncated to MAX_COMPUTERNAME_LENGTH characters, not
        '       including the terminating null character.
        '   cnfComputerNamePhysicalDnsDomain
        '       Sets the name of the DNS domain assigned to the computer.
        Select Case NameFormat
            Case cnfComputerNamePhysicalNetbios
                If Len(Name) > dhcMaxComputerName Then
                    With Err
                        .Raise ERR_INVALID_NAME, _
                        "SystemInfo.ComputerNameEx", _
                         "Name cannot exceed " & _
                         dhcMaxComputerName & " characters."
                    End With
                End If
            Case cnfComputerNamePhysicalDnsHostname
                If Len(Name) > dhcMaxComputerName Then
                    Call HandleErrors(ERR_NAME_TOO_LONG, _
                     "NetBIOS name is longer than " & _
                     dhcMaxComputerName & " characters.")
                End If
            Case cnfComputerNamePhysicalDnsDomain
                ' It's here just so that we can escape the Else clause.
            Case Else
                ' For Public Property Let, only the above three
                ' values are acceptable.
                Err.Raise 5
        End Select
        Call SetComputerNameEx(NameFormat, Name)
    Else
        ' Either the OS is not Win2000 or NameFormat
        ' is 0 or cnfComputerNameNetBIOS, so use the
        ' normal API functions
        If NameFormat = cnfComputerNameNetBIOS Then
            Call SetComputerName(Name)
        Else
            If RaiseErrors Then
                Call HandleErrors(ERR_INVALID_OS)
            End If
        End If
    End If
End Property

Public Property Get UserName( _
 Optional ExtendedFormat As ExtendedNameFormat = enfNameUnknown) _
 As String

    ' Retrieve the name of the logged-in user.
    ' On Windows 2000, retrieves the name of the user or other security
    ' principal associated with the calling thread.
    '
    ' It appears that GetUserName counts the trailing null in the length it
    ' places in lngLen.

    Dim lngLen As Long
    Dim strBuffer As String
    Dim lngRet As Long

    Const dhcMaxUserName = 255

    ' Initialize the buffer strings
    strBuffer = String$(dhcMaxUserName, vbNullChar)
    lngLen = dhcMaxUserName
    If IsWin2000 Then
        If ExtendedFormat <> enfNameUnknown Then
            ' If a particular ExtendedFormat is requested and the
            ' OS is Windows 2000, then use the Extended version
            ' of the API function.
            lngRet = GetUserNameEx(ExtendedFormat, strBuffer, lngLen)
            ' Even if lngRet and Err.LastDLLError indicate that
            ' the call to GetUserNameEx was successful,
            ' strBuffer and lngLen may not get modified, in which case
            ' strBuffer will still contain only vbNullChars. To make
            ' sure that a valid string was returned in strBuffer,
            ' check lngRet and the length of strBuffer
            ' after trimming to the first instance of vbNullChar
            If lngRet And Len(dhTrimNull(strBuffer)) > 0 Then
                ' If successful, return the username
                UserName = Left$(strBuffer, lngLen - 1)
            Else
                If RaiseErrors Then
                    With Err
                        .Raise .LastDllError, _
                         "SystemInfo.UserName", APIErr(.LastDllError)
                    End With
                End If
            End If
        Else
            ' Specified ExtendedFormat was enfNameUnknown
            ' use GetUserName instead
            If CBool(GetUserName(strBuffer, lngLen)) Then
                UserName = Left$(strBuffer, lngLen - 1)
            End If
        End If
    Else
        ' OS is not Win2000
        ' In this case, only enfNameUnknown is valid
        If ExtendedFormat = enfNameUnknown Then
            ' use GetUserName API function
            If CBool(GetUserName(strBuffer, lngLen)) Then
                UserName = Left$(strBuffer, lngLen - 1)
            End If
        Else
            If RaiseErrors Then
                Call HandleErrors(ERR_INVALID_OS)
            End If
        End If
    End If
End Property

Public Property Get WINVER() As Long
    ' Equivalent to SDK's WINVER environment variable
    
    If IsWin95 Or IsWinNT Then
        WINVER = 4&
    End If
    If IsWin98 Or IsWin2000 Then
        WINVER = 5&
    End If
End Property

Public Property Get WIN32_IE() As Long
    ' Equivalent to SDK's _WIN32_IE environment variable
    
    '
    ' The Major and Minor values of the dll together
    ' constitute the Version Public Property.
    Const IE_DLL = "ShDocVW.dll"        '  The core IE DLL in System32 folder
    Dim strFileVersionMajor As String
    Dim strFileVersionMinor As String
        
    ' Fill in the ByRef Major and Minor arguments with the file's version
    Call GetFileVersion(IE_DLL, strFileVersionMajor, strFileVersionMinor)
    
    ' Now set the return value based on the table above
    Select Case strFileVersionMajor
        Case "4.70"
            WIN32_IE = 3&
        Case "4.71", "4.72"
            WIN32_IE = 4&
        Case "5.0"
            WIN32_IE = 5&
    End Select
    If WINVER = 5& Then WIN32_IE = 4&
End Property

Public Property Get WIN32_WINDOWS() As Long
    ' Equivalent to SDK's _WIN32_WINDOWS environment variable
    
    If IsWin98 Then
        WIN32_WINDOWS = 410&
    ElseIf IsWin95 Then
        WIN32_WINDOWS = 4&
    End If
End Property

Public Property Get WIN32_WINNT() As Long
    ' Equivalent to SDK's _WIN32_WINNT environment variable
    
    If IsWinNT Then
        WIN32_WINNT = 4&
    End If
    If IsWin2000 Then
        WIN32_WINNT = 5&
    End If
End Property

Public Property Get ServicePackMajorVersion() As Integer
    ' Returns the major version number of the latest
    ' Service Pack installed on the system, Win2000
    
    If IsWin2000 And mblnVersionInfoEx Then
        ServicePackMajorVersion = osvi.wServicePackMajor
    Else
        Call HandleErrors(ERR_INVALID_OS)
    End If
End Property

Public Property Get ServicePackMinorVersion() As Integer
    ' Returns the minor version number of the latest
    ' Service Pack installed on the system, Win2000
    
    If IsWin2000 And mblnVersionInfoEx Then
        ServicePackMinorVersion = osvi.wServicePackMinor
    Else
        Call HandleErrors(ERR_INVALID_OS)
    End If
End Property

Public Property Get IsSuiteInstalled(SuiteType As ProductSuiteType) As Boolean
    ' Returns true if the specified suite is installed on Win2000
    
    If IsWin2000 And mblnVersionInfoEx Then
        If osvi.wSuiteMask And SuiteType Then
            IsSuiteInstalled = True
        End If
    Else
        Call HandleErrors(ERR_INVALID_OS)
    End If
End Property

Public Property Get OSMajorVersion() As Long
    ' Retrieve the major version number of the operating system.
    ' For example, for Windows NT version 3.51, the major version
    ' number is 3; and for Windows NT version 4.0, the major version
    ' number is 4.
    OSMajorVersion = osvi.dwMajorVersion
End Property

Public Property Get OSMinorVersion() As Long
    ' Retrieve the minor version number of the operating system.
    ' For example, for Windows NT version 3.51, the minor version
    ' number is 51; and for Windows NT version 4.0, the minor
    ' version number is 0.
    OSMinorVersion = osvi.dwMinorVersion
End Property

Public Property Get OSBuild() As Long
    ' Retrieve the build number of the operating system.
    If IsWin95 Then
        OSBuild = osvi.dwBuildNumber And &HFF
    Else
        OSBuild = osvi.dwBuildNumber
    End If
End Property

Public Property Get OSVersion() As String
    ' Builds a string with OS Description, like
    ' "Microsoft Windows NT Server version 4.0 Service Pack 4 (Build 1381)"
    
    Dim strOut As String
    Dim hKey As Long
    Dim szProductType As String
    Dim dwBufLen As Long

    If IsWinNT Or IsWin2000 Then
        Select Case OSMajorVersion
            Case Is <= 4
                strOut = "Microsoft Windows NT "
            Case 5
                strOut = "Microsoft Windows 2000 "
        End Select
        If mblnVersionInfoEx Then
            ' if OSVERSIONINFOEX UDT was used when calling GetVersionEx,
            ' then the ProductType info is already available as a member of the UDT.
            strOut = strOut & (ProductType + Chr$(vbKeySpace))
        Else
            ' if OSVERSIONINFO was used, then we have to read the ProductType
            ' information from the registry.
            Call RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
             "SYSTEM\CurrentControlSet\Control\ProductOptions", _
             0, KEY_QUERY_VALUE, hKey)
            
            ' if the registry open operation was successful, continue onwards
            If hKey Then
                szProductType = Space$(80)
                dwBufLen = Len(szProductType)
                ' Read the value from Registry and close the key
                Call RegQueryValueEx(hKey, "ProductType", 0, 0, ByVal szProductType, dwBufLen)
                Call RegCloseKey(hKey)
                szProductType = Left$(szProductType, dwBufLen - 1)
                ' if szProductType is same as "WINNT" then the OS is NT Workstation
                If StrComp(szProductType, "WINNT", vbTextCompare) = 0 Then
                    strOut = strOut & "Workstation "
                ' otherwise if szProductType is "SERVERNT", then the OS is NT Server
                ElseIf StrComp(szProductType, "SERVERNT", vbTextCompare) = 0 Then
                    strOut = strOut & "Server "
                End If
            End If
        End If
        
        ' build the complete string for WinNT or Win2000
        strOut = strOut & _
         "version " & OSMajorVersion & "." & OSMinorVersion & _
         " " & OSExtraInfo & " (Build " & OSBuild & ")"
    ElseIf IsWin95 Then
        ' Nothing special for Win95 and Win98
        strOut = "Microsoft Windows 95"
    ElseIf IsWin98 Then
        strOut = "Microsoft Windows 98"
    End If
    ' return the string
    OSVersion = strOut
End Property

Public Property Get OSExtraInfo() As String
    ' Retrieve extra operating system information, like "Service Pack 3".
     OSExtraInfo = dhTrimNull(osvi.szCSDVersion)
End Property

Public Property Get ProductType() As String
    ' Returns the "make" of Win2000, like "Professional"
    
    If IsWin2000 Then
        Select Case osvi.wProductType
            Case VER_NT_WORKSTATION
                ProductType = "Professional"
            Case VER_NT_SERVER
                ProductType = "Server"
            Case VER_NT_DOMAIN_CONTROLLER
                ProductType = "Domain Controller"
        End Select
    Else
        Call HandleErrors(ERR_INVALID_OS)
    End If
End Property

Public Property Get IsWin95() As Boolean
    ' Returns True if the operating system is Windows 95.
    
    With osvi
        IsWin95 = (.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS _
         And .dwMinorVersion = 0)
    End With
End Property

Public Property Get IsWin98() As Boolean
    ' Returns True if the operating system is Windows 98.
    
    With osvi
        IsWin98 = (.dwMajorVersion = 4 And _
         (.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS _
         And .dwMinorVersion > 0))
    End With
End Property

Public Property Get IsWin2000() As Boolean
    ' Returns True if the operating system is Windows 2000.
    
    With osvi
        IsWin2000 = (.dwPlatformId = VER_PLATFORM_WIN32_NT _
         And .dwMajorVersion = 5)
    End With
End Property

Public Property Get IsWinNT() As Boolean
    ' Returns True if the operating system is Windows NT.
    
    With osvi
        IsWinNT = (.dwPlatformId = VER_PLATFORM_WIN32_NT _
         And .dwMajorVersion <= 4)
    End With
End Property

Public Property Get ProcessorArchitecture() As ProcessorType
    ' Specifies the system's processor architecture.
    ' One of:
    ' 0     Intel   (WinNT or Win95)
    ' 1     MIPS    (WinNT only)
    ' 2     ALPHA   (WinNT only)
    ' 3     PPC     (WinNT only)
    ' -1    Unknown (WinNT only)
    ProcessorArchitecture = si.wProcessorArchitecture
End Property

Public Property Get PageSize() As Long
    ' Specifies the page size and the granularity of
    ' page protection and commitment.
    PageSize = si.dwPageSize
End Property

Public Property Get MinAppAddress() As Long
    ' Pointer to the lowest memory address accessible to
    ' applications and dynamic-link libraries (DLLs).
    MinAppAddress = si.lpMinimumApplicationAddress
End Property

Public Property Get MaxAppAddress() As Long
    ' Pointer to the highest memory address accessible to
    ' applications and DLLs.
    MaxAppAddress = si.lpMaximumApplicationAddress
End Property

Public Property Get ActiveProcessorMask() As Long
    ' Specifies a mask representing the set of
    ' processors configured into the system.
    ' Bit 0 is processor 0; bit 31 is processor 31.
    ActiveProcessorMask = si.dwActiveProcessorMask
End Property

Public Property Get NumberOfProcessors() As Long
    ' Specifies the number of processors in the system.
    NumberOfProcessors = si.dwNumberOrfProcessors
End Property

Public Property Get ProcessorType() As Long
    ' Win95:Specifies the type of processor in the system.
    ' WinNT: use ProcessorArchitecture, ProcessorLevel,
    ' and ProcessorRevision values.
    ProcessorType = si.dwProcessorType
End Property

Public Property Get AllocationGranularity() As Long
    ' Specifies the granularity with which virtual memory is allocated.
    AllocationGranularity = si.dwAllocationGranularity
End Property

Public Property Get ProcessorLevel() As Integer
    ' Windows 95: Not used.
    ' Windows NT: Specifies the system's architecture-dependent processor level.
    ' For Intel:
    ' 3  Intel 80386
    ' 4  Intel 80486
    ' 5  Pentium
    ' 6  Intel Pentium Pro or Pentium II
    ' For other processors, see MSDN or other documentation.
    ProcessorLevel = si.wProcessorLevel
End Property

Public Property Get ProcessorRevision() As Integer
    ' Windows 95: This member is not used.
    ' Windows NT: Specifies an architecture-dependent processor revision.
    ProcessorRevision = si.wProcessorRevision
End Property

Public Property Get SpecialFolderLocation(ByVal CSIDL As siCSIDL_VALUES) As String
    ' Returns path to a special folder on the machine
    ' without a trailing backslash.
    Dim lngRet As Long
    Dim strLocation As String
    Dim pidl As Long

    ' retrieve a PIDL for the specified location
    lngRet = SHGetSpecialFolderLocation(0, CSIDL, pidl)
    If lngRet = NOERROR Then
        strLocation = Space$(MAX_PATH)
        '  convert the pidl to a physical path
        lngRet = SHGetPathFromIDList(ByVal pidl, strLocation)
        If lngRet Then
            ' if successful, return the location
            SpecialFolderLocation = dhTrimNull(strLocation)
        End If
        ' Freeup the allocatted memory
        Call CoTaskMemFree(pidl)
    End If
End Property

Private Function dhTrimNull(strValue As String) As String
    ' Borrowed from Chapter 1.
    Dim intPos As Integer
    intPos = InStr(strValue, vbNullChar)
    Select Case intPos
        Case 0
            dhTrimNull = strValue
        Case 1
            dhTrimNull = ""
        Case Else
            dhTrimNull = Left$(strValue, intPos - 1)
    End Select
End Function

Private Sub GetFileVersion(ByVal strFile As String, _
 strFileVersionMS As String, _
 strFileVersionLS As String)
    
    ' For executable files (Applications),
    ' return the version number.
    '
    On Error GoTo ErrHandler
    Dim lngSize As Long
    Dim lngRet As Long
    Dim pBlock() As Byte
    Dim lpfi As VS_FIXEDFILEINFO
    Dim lppBlock As Long

    ' GetFileVersionInfo requires us to get the size
    ' of the file version information first, this info is in the format
    ' of VS_FIXEDFILEINFO struct
    lngSize = GetFileVersionInfoSize(strFile, lngRet)

    ' If the OS can obtain version info, then proceed on
    If lngSize Then
        ' The info in pBlock is always in Unicode format
        ReDim pBlock(lngSize)
        lngRet = GetFileVersionInfo(strFile, 0, lngSize, pBlock(0))
        If lngRet Then
            ' The same pointer to pBlock can be passed to VerQueryValue
            lngRet = VerQueryValue(pBlock(0), "\", lppBlock, lngSize)
            
            ' Fill the VS_FIXEDFILEINFO struct with bytes from pBlock
            ' VerQueryValue fills lngSize with the length of the block.
            Call CopyMemory(lpfi, ByVal lppBlock, lngSize)
            ' Build the version info strings
            strFileVersionMS = CStr(HIWord(lpfi.dwFileVersionMS)) & _
             "." & CStr(LOWord(lpfi.dwFileVersionMS))
            strFileVersionLS = CStr(HIWord(lpfi.dwFileVersionLS)) & _
             "." & CStr(LOWord(lpfi.dwFileVersionLS))
       End If
    End If

ExitHere:
    Erase pBlock
    Exit Sub
    
ErrHandler:
    Resume ExitHere
End Sub

Private Function APIErr(ByVal lngErr As Long) As String
    
    ' retrieves the error description for specific values of LastDllError
    Dim strMsg As String
    Dim lngRet As Long
    
    strMsg = String$(1024, 0)
    lngRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, _
     lngErr, 0, strMsg, Len(strMsg), ByVal 0&)
    If lngRet Then APIErr = Left$(strMsg, lngRet)
End Function

Private Function LOWord(dw As Long) As Integer
    
    '  retrieves the low-order word from the given 32-bit value.
    If dw And &H8000& Then
        LOWord = dw Or &HFFFF0000
    Else
        LOWord = dw And &HFFFF&
    End If
End Function

Private Function HIWord(dw As Long) As Integer
    
    '  retrieves the high-order word from the given 32-bit value.
  HIWord = (dw And &HFFFF0000) \ &H10000
End Function

Private Sub HandleErrors( _
 lngErrCode As Long, _
 Optional strErrMsg As String)
    ' Centralized error handler to raise
    ' the errors to the client
    
    With Err
        If RaiseErrors Then
            If Len(strErrMsg) > 0 Then
                .Raise .Number, "SystemInfo", .Description, .HelpFile, .HelpContext
            Else
                .Raise lngErrCode, "SystemInfo", ERR_STRING
            End If
        End If
    End With
End Sub

Private Sub Class_Initialize()
    Dim osviTmp As OSVERSIONINFO

    ' Set the flag to true so that an error is raised
    ' if a non-applicable Public Property is used for a particular
    ' operating system
    RaiseErrors = True

    ' First try with OSVersionInfoEx
    osvi.dwOSVersionInfoSize = Len(osvi)
    mblnVersionInfoEx = CBool(GetVersionEx(osvi))
    If Not mblnVersionInfoEx Then
        ' If it failed, then you aren't running Win2000
        ' so try with OSVersionInfo.
        ' Changing the Size member tells the OS
        ' which UDT you want the info for.
        osvi.dwOSVersionInfoSize = Len(osviTmp)
        Call GetVersionEx(osvi)
    End If
    ' Get the other information as well
    Call GetSystemInfo(si)
End Sub





