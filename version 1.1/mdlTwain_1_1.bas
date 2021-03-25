Attribute VB_Name = "mdlTwain"
'*******************************************************************************
'
' Description: VB Module for accessing TWAIN compatible scanner (VB 5, 6)
'
' Author:      Lumir Mik (lumir.mik@gmail.com)
'
' Version:     1.1
'
' License:     Free to any use. If you change some part of this code, please,
'              mention it here.
'              Consider this code my contribution to free programmer sources
'              in which I found much help and inspiration.
'
' There are 3 public functions in this module:
'
'   1. PopupSelectSourceDialog
'           shows TWAIN dialog for selecting default source for acquisition
'
'   2. TransferWithoutUI
'           transfers image(s) from TWAIN data source without showing
'           the data source user interface (silent transfer). The programmer
'           can set following attributes of the image:
'               - resolution (DPI)
'               - colour depth - monochromatic, grey, fullcolour
'               - image size and position
'                       - left, top, right, bottom (in inches).
'           The programmer can also activate ADF (Automatic Document Feeder).
'           The image is saved into the BMP file.
'
'   3. TransferWithUI
'           transfers image(s) from TWAIN data source using the data
'           source user interface to set image and transfer attributes.
'           The image is saved into the BMP file.
'
'*******************************************************************************

Option Explicit


'------------------------------
' Declaration for TWAIN_32.DLL
'------------------------------
Private Declare Function DSM_Entry Lib "Twain_32.dll" ( _
                                   ByRef pOrigin As Any, _
                                   ByRef pDest As Any, _
                                   ByVal DG As Long, _
                                   ByVal DAT As Integer, _
                                   ByVal MSG As Integer, _
                                   ByRef pData As Any) As Integer

Private Type TW_VERSION
    MajorNum       As Integer                    ' TW_UINT16
    MinorNum       As Integer                    ' TW_UINT16
    Language       As Integer                    ' TW_UINT16
    Country        As Integer                    ' TW_UINT16
    Info(1 To 34)  As Byte                       ' TW_STR32
End Type

Private Type TW_IDENTITY
    Id                      As Long              ' TW_UINT32
    Version                 As TW_VERSION        ' TW_VERSION
    ProtocolMajor           As Integer           ' TW_UINT16
    ProtocolMinor           As Integer           ' TW_UINT16
    SupportedGroups1        As Integer           ' TW_UINT32
    SupportedGroups2        As Integer
    Manufacturer(1 To 34)   As Byte              ' TW_STR32
    ProductFamily(1 To 34)  As Byte              ' TW_STR32
    ProductName(1 To 34)    As Byte              ' TW_STR32
End Type

Private Type TW_USERINTERFACE
    ShowUI   As Integer                          ' TW_BOOL
    ModalUI  As Integer                          ' TW_BOOL
    hParent  As Long                             ' TW_HANDLE
End Type

Private Type TW_PENDINGXFERS
    Count     As Integer                         ' TW_UINT16
    Reserved  As Long                            ' TW_UINT32
End Type

Private Type TW_ONEVALUE
    ItemType  As Integer                         ' TW_UINT16
    Item      As Long                            ' TW_UINT32
End Type

Private Type TW_CAPABILITY
    Cap         As Integer                       ' TW_UINT16
    ConType     As Integer                       ' TW_UINT16
    hContainer  As Long                          ' TW_HANDLE
End Type

Private Type TW_FIX32
    Whole  As Integer                            ' TW_INT16
    Frac   As Integer                            ' TW_UINT16
End Type

Private Type TW_FRAME
    Left    As TW_FIX32                          ' TW_FIX32
    Top     As TW_FIX32                          ' TW_FIX32
    Right   As TW_FIX32                          ' TW_FIX32
    Bottom  As TW_FIX32                          ' TW_FIX32
End Type

Private Type TW_IMAGELAYOUT
    Frame           As TW_FRAME                  ' TW_FRAME
    DocumentNumber  As Long                      ' TW_UINT32
    PageNumber      As Long                      ' TW_UINT32
    FrameNumber     As Long                      ' TW_UINT32
End Type

Private Type TW_EVENT
    pEvent     As Long                           ' TW_MEMREF
    TWMessage  As Integer                        ' TW_UINT16
End Type


Private Const DG_CONTROL = 1
Private Const DG_IMAGE = 2

Private Const MSG_GET = 1
Private Const MSG_SET = 6
Private Const MSG_RESET = 7
Private Const MSG_XFERREADY = 257
Private Const MSG_CLOSEDSREQ = 258
Private Const MSG_OPENDSM = 769
Private Const MSG_CLOSEDSM = 770
Private Const MSG_OPENDS = 1025
Private Const MSG_CLOSEDS = 1026
Private Const MSG_USERSELECT = 1027
Private Const MSG_DISABLEDS = 1281
Private Const MSG_ENABLEDS = 1282
Private Const MSG_PROCESSEVENT = 1537
Private Const MSG_ENDXFER = 1793

Private Const DAT_CAPABILITY = 1
Private Const DAT_EVENT = 2
Private Const DAT_IDENTITY = 3
Private Const DAT_PARENT = 4
Private Const DAT_PENDINGXFERS = 5
Private Const DAT_USERINTERFACE = 9
Private Const DAT_IMAGELAYOUT = 258
Private Const DAT_IMAGENATIVEXFER = 260

Private Const TWRC_SUCCESS = 0
Private Const TWRC_FAILURE = 1
Private Const TWRC_CHECKSTATUS = 2
Private Const TWRC_CANCEL = 3
Private Const TWRC_NOTDSEVENT = 5
Private Const TWRC_XFERDONE = 6

Private Const TWLG_CZECH = 45

Private Const TWCY_CZECHREPUBLIC = 420

Private Const TWON_PROTOCOLMAJOR = 1
Private Const TWON_ONEVALUE = 5
Private Const TWON_PROTOCOLMINOR = 9


'--------------------------
' Declaration for WIN32API
'--------------------------
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
                               ByVal pDest As Long, _
                               ByVal pSource As Long, _
                               ByVal Length As Long)

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" ( _
                               ByVal pDest As Long, _
                               ByVal Length As Long)

Private Declare Function GlobalFree Lib "kernel32.dll" ( _
                                    ByVal hMem As Long) As Long

Private Declare Function GlobalLock Lib "kernel32.dll" ( _
                                    ByVal hMem As Long) As Long

Private Declare Function GlobalUnlock Lib "kernel32.dll" ( _
                                      ByVal hMem As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32.dll" ( _
                                     ByVal wFlags As Long, _
                                     ByVal dwBytes As Long) As Long

Private Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" ( _
                                    ByRef lpMsg As MSG, _
                                    ByVal hWnd As Long, _
                                    ByVal wMsgFilterMin As Long, _
                                    ByVal wMsgFilterMax As Long) As Long

Private Declare Function TranslateMessage Lib "user32.dll" ( _
                                          ByRef lpMsg As MSG) As Long

Private Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" ( _
                                         ByRef lpMsg As MSG) As Long

Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" ( _
                                        ByVal dwExStyle As Long, _
                                        ByVal lpClassName As String, _
                                        ByVal lpWindowName As String, _
                                        ByVal dwStyle As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal nWidth As Long, _
                                        ByVal nHeight As Long, _
                                        ByVal hWndParent As Long, _
                                        ByVal hMenu As Long, _
                                        ByVal hInstance As Long, _
                                        ByVal lpParam As Long) As Long

Private Declare Function DestroyWindow Lib "user32.dll" ( _
                                       ByVal hWnd As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" ( _
                                     ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32.dll" ( _
                                     ByVal hLibModule As Long) As Long


Private Type BITMAPFILEHEADER
    bfType       As Integer
    bfSize       As Long
    bfReserved1  As Integer
    bfReserved2  As Integer
    bfOffBits    As Long
End Type

Private Type BITMAPINFOHEADER
    biSize           As Long
    biWidth          As Long
    biHeight         As Long
    biPlanes         As Integer
    biBitCount       As Integer
    biCompression    As Long
    biSizeImage      As Long
    biXPelsPerMeter  As Long
    biYPelsPerMeter  As Long
    biClrUsed        As Long
    biClrImportant   As Long
End Type

Private Type RGBQUAD
    rgbBlue      As Byte
    rgbGreen     As Byte
    rgbRed       As Byte
    rgbReserved  As Byte
End Type

Private Type POINTAPI
    X  As Long
    Y  As Long
End Type

Private Type MSG
    hWnd     As Long
    Message  As Long
    wParam   As Long
    lParam   As Long
    time     As Long
    pt       As POINTAPI
End Type


Private Const GHND = 66


'-----------------------------
' Declaration for this Module
'-----------------------------
Private m_tAppID As TW_IDENTITY
Private m_tSrcID As TW_IDENTITY
Private m_lHndMsgWin As Long


Public Enum TWAIN_MDL_COLOURTYPE
    BW = 0                                       ' TWPT_BW
    GREY = 1                                     ' TWPT_GRAY
    RGB = 2                                      ' TWPT_RGB
End Enum

Private Enum TWAIN_MDL_CAPABILITY
    XFERCOUNT = 1                                ' CAP_XFERCOUNT
    PIXELTYPE = 257                              ' ICAP_PIXELTYPE
    FEEDERENABLED = 4098                         ' CAP_FEEDERENABLED
    INDICATORS = 4107                            ' CAP_INDICATORS
    PHYSICALWIDTH = 4369                         ' ICAP_PSYSICALWIDTH
    PHYSICALHEIGHT = 4370                        ' ICAP_PSYSICALHEIGHT
    XRESOLUTION = 4376                           ' ICAP_XRESOLUTION
    YRESOLUTION = 4377                           ' ICAP_YRESOLUTION
    BITDEPTH = 4395                              ' ICAP_BITDEPTH
End Enum

Private Enum TWAIN_MDL_ITEMYPE
    INT16 = 1                                    ' TW_INT16  (short)
    UINT16 = 4                                   ' TW_UINT16 (unsigned short)
    BOOL = 6                                     ' TW_BOOL   (unsigned short)
    FIX32 = 7                                    ' TW_FIX32  (structure)
End Enum


Public Function TransferWithoutUI(ByVal sngResolution As Single, _
                                  ByVal iColourType As TWAIN_MDL_COLOURTYPE, _
                                  ByVal sngImageLeft As Single, _
                                  ByVal sngImageTop As Single, _
                                  ByVal sngImageRight As Single, _
                                  ByVal sngImageBottom As Single, _
                                  ByVal sBMPFileNameWithoutExt As String, _
                                  ByRef lFileCounter As Long, _
                                  ByVal blADF As Boolean) As Long
    '-----------------------------------------------------------------------------------------
    ' Function transfers image(s) from Twain data source without showing
    '   the data source user interface (silent transfer).
    '
    ' Filename of the saved image(s) is in the format
    '   sBMPFileNameWithoutExt_iFileCounter.bmp, e.g. Test_000001.bmp
    '
    ' Input values
    '   - sngResolution (Single) - resolution of the image in DPI (dots per inch)
    '   - iColourType (Integer)  - colour depth of the image - monochromatic (BW),
    '                              colours of grey (GREY), full colours (COLOUR)
    '   - sngImageLeft, sngImageTop, sngImageRight, sngImageBottom (Single) -
    '       values determine the rectangle that will be scanned (default units are inches).
    '       If you set Right and Bottom values to 0, the module sets maximum values
    '       allowed by the scanner driver (e.g. the bottom right corner of the scanner glass)
    '   - sBMPFileNameWithoutExt (String) - the file name of the saved image(s),
    '       without extension .bmp
    '   - lFileCounter (Long) - counter, which is put at the end of image file name.
    '       This value is passed by reference and is incremented with each
    '       successful transfer / saving. It is formatted as 6 digits number
    '       with leading zeros (000000).
    '   - blADF (Boolean) - if TRUE, the Automatic Document Feeder is activated.
    '       If FALSE, the scanner glass is scanned, even though ADF is present
    '
    ' Function returns 0 if OK, 1 if ERROR
    '-----------------------------------------------------------------------------------------

    Dim lRtn As Long

    On Local Error GoTo ErrPlace

    
    '----------------------
    ' FileCounter checking
    '----------------------
    If (lFileCounter < 1) Or (lFileCounter > 999999) Then GoTo ErrPlace
    

    '---------------------------
    ' Try to load Twain library
    '---------------------------
    lRtn = LoadTwainDll()
    If lRtn Then GoTo ErrPlace

    '--------------------------------
    ' Open Twain Data Source Manager
    '--------------------------------
    lRtn = OpenTwainDSM()
    If lRtn Then GoTo ErrPlace

    '------------------------
    ' Open Twain Data Source
    '------------------------
    lRtn = OpenTwainDS()
    If lRtn Then GoTo ErrPlace

    '---------------------------------------------------------------------------
    ' SET ALL IMPORTANT IMAGE AND TRANSFER ATTRIBUTES
    '
    ' Activate ADF (Automatic Document Feeder) or just scan the glass
    '   Some scanners do not support this feature, so do not check return value
    '---------------------------------------------------------------------------
    If blADF Then
        lRtn = TwainSetOnevalue(FEEDERENABLED, BOOL, 1)
        'If lRtn Then GoTo ErrPlace
    Else
        lRtn = TwainSetOnevalue(FEEDERENABLED, BOOL, 0)
        'If lRtn Then GoTo ErrPlace
    End If

    '--------------------------------------------------------
    ' Set image size and position
    ' If sngImageRight = 0, replace it with physical width
    ' if sngImageBottom = 0, replace it with physical height
    '--------------------------------------------------------
    If sngImageRight = 0 Then
        lRtn = TwainGetOnevalue(PHYSICALWIDTH, sngImageRight)
        If lRtn Then GoTo ErrPlace
    End If
    If sngImageBottom = 0 Then
        lRtn = TwainGetOnevalue(PHYSICALHEIGHT, sngImageBottom)
        If lRtn Then GoTo ErrPlace
    End If
    lRtn = SetImageSize(sngImageLeft, sngImageTop, sngImageRight, sngImageBottom)
    If lRtn Then GoTo ErrPlace

    '------------------------------------------------
    ' Set the image resolution in DPI - both X and Y
    '------------------------------------------------
    lRtn = TwainSetOnevalue(XRESOLUTION, FIX32, sngResolution)
    If lRtn Then GoTo ErrPlace

    lRtn = TwainSetOnevalue(YRESOLUTION, FIX32, sngResolution)
    If lRtn Then GoTo ErrPlace

    '---------------------------
    ' Set the image colour type
    '---------------------------
    lRtn = TwainSetOnevalue(PIXELTYPE, UINT16, iColourType)
    If lRtn Then GoTo ErrPlace

    '-----------------------------------------------------------------
    ' If the colour type is fullcolour, set the bitdepth of the image
    '   - 24 bits, 32 bits, ...
    '-----------------------------------------------------------------
    If iColourType = RGB Then lRtn = TwainSetOnevalue(BITDEPTH, UINT16, 24)

    '-------------------------------------
    ' TRANSFER the image with UI disabled
    '-------------------------------------
    lRtn = TwainTransfer(False, sBMPFileNameWithoutExt, lFileCounter)
    If lRtn Then GoTo ErrPlace

    '-------------------
    ' Close Data Source
    '-------------------
    lRtn = CloseTwainDS()
    If lRtn Then GoTo ErrPlace

    '---------------------------
    ' Close Data Source Manager
    '---------------------------
    lRtn = CloseTwainDSM()
    If lRtn Then GoTo ErrPlace

    TransferWithoutUI = 0
    Exit Function

ErrPlace:
    lRtn = CloseTwainDS()
    lRtn = CloseTwainDSM()
    TransferWithoutUI = 1
End Function

Public Function TransferWithUI(ByVal sBMPFileNameWithoutExt As String, _
                               ByRef lFileCounter As Long) As Long
    '-----------------------------------------------------------------------------------
    ' Function transfers image(s) from Twain data source using the data
    '   source user interface to set image and transfer attributes.
    '
    ' Filename of the saved image(s) is in the format
    '   sBMPFileNameWithoutExt_iFileCounter.bmp, e.g. Test_000001.bmp
    '
    ' Input values
    '   - sBMPFileNameWithoutExt (String) - the file name of the saved image(s),
    '       without extension .bmp
    '   - lFileCounter (Long) - counter, which is put at the end of image file name.
    '       This value is passed by reference and is incremented with each
    '       successful transfer / saving. It is formatted as 6 digits number
    '       with leading zeros (000000).
    '
    ' Function returns 0 if OK, 1 if an error occurs
    '-----------------------------------------------------------------------------------

    Dim lRtn As Long

    On Local Error GoTo ErrPlace


    '----------------------
    ' FileCounter checking
    '----------------------
    If (lFileCounter < 1) Or (lFileCounter > 999999) Then GoTo ErrPlace

    
    '---------------------------
    ' Try to load Twain library
    '---------------------------
    lRtn = LoadTwainDll()
    If lRtn Then GoTo ErrPlace

    '--------------------------------
    ' Open Twain Data Source Manager
    '--------------------------------
    lRtn = OpenTwainDSM()
    If lRtn Then GoTo ErrPlace

    '------------------------
    ' Open Twain Data Source
    '------------------------
    lRtn = OpenTwainDS()
    If lRtn Then GoTo ErrPlace

    '------------------------------------
    ' TRANSFER the image with UI enabled
    '------------------------------------
    lRtn = TwainTransfer(True, sBMPFileNameWithoutExt, lFileCounter)
    If lRtn Then GoTo ErrPlace

    '-------------------
    ' Close Data Source
    '-------------------
    lRtn = CloseTwainDS()
    If lRtn Then GoTo ErrPlace

    '---------------------------
    ' Close Data Source Manager
    '---------------------------
    lRtn = CloseTwainDSM()
    If lRtn Then GoTo ErrPlace

    TransferWithUI = 0
    Exit Function

ErrPlace:
    lRtn = CloseTwainDS()
    lRtn = CloseTwainDSM()
    TransferWithUI = 1
End Function

Public Function PopupSelectSourceDialog() As Long
    '-------------------------------------------------------------------
    ' Function shows the Twain dialog for selecting default data source
    ' Function returns 0 if OK, 1 if an error occurs
    '-------------------------------------------------------------------

    Dim lRtn As Long
    Dim iRtn As Integer

    On Local Error GoTo ErrPlace


    '---------------------------
    ' Try to load Twain library
    '---------------------------
    lRtn = LoadTwainDll()
    If lRtn Then GoTo ErrPlace

    '--------------------------------
    ' Open Twain Data Source Manager
    '--------------------------------
    lRtn = OpenTwainDSM()
    If lRtn Then GoTo ErrPlace

    '--------------------------------------------
    ' Popup "Select source" dialog
    '   DG_CONTROL, DAT_IDENTITY, MSG_USERSELECT
    '--------------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, MSG_USERSELECT, m_tSrcID)
    If iRtn <> TWRC_SUCCESS Then
        lRtn = CloseTwainDSM()
        GoTo ErrPlace
    End If

    '---------------------------------
    ' Close Twain Data Source Manager
    '---------------------------------
    lRtn = CloseTwainDSM()
    If lRtn Then GoTo ErrPlace

    PopupSelectSourceDialog = 0
    Exit Function

ErrPlace:
    PopupSelectSourceDialog = 1
End Function

Private Function LoadTwainDll() As Long
    '--------------------------------------------------
    ' Function tries to load TWAIN_32.DLL library file
    ' Function returns 0 if OK, 1 if ERROR
    '--------------------------------------------------

    Dim lhLib As Long

    On Local Error GoTo ErrPlace


    lhLib = LoadLibrary("TWAIN_32.DLL")
    If lhLib = 0 Then GoTo ErrPlace

    Call FreeLibrary(lhLib)

    LoadTwainDll = 0
    Exit Function

ErrPlace:
    LoadTwainDll = 1
End Function

Private Function OpenTwainDSM() As Long
    '---------------------------------------
    ' Function opens the Data Source Manger
    ' Function returns 0 if OK, 1 if ERROR
    '---------------------------------------

    Dim sTmp As String
    Dim iRtn As Integer

    On Local Error GoTo ErrPlace


    '-----------------------------------------------------
    ' Create window that will receive all TWAIN messages
    ' Message loop can be found in TwainTransfer function
    '-----------------------------------------------------
    m_lHndMsgWin = CreateWindowEx(0&, "#32770", "TWAIN_MSG_WINDOW", 0&, 10&, 10&, 150&, 50&, 0&, 0&, 0&, 0&)
    If m_lHndMsgWin = 0 Then GoTo ErrPlace

    '-----------------------------------------------------------------------------------------
    ' Introduce yourself to TWAIN
    ' - MajorNum, MinorNum, Language, Country, Manufacturer, ProductFamily, ProductName, etc.
    '-----------------------------------------------------------------------------------------
    Call ZeroMemory(VarPtr(m_tAppID), Len(m_tAppID))
    With m_tAppID
        .Version.MajorNum = 1
        .Version.MinorNum = 1
        .Version.Language = TWLG_CZECH
        .Version.Country = TWCY_CZECHREPUBLIC
        .ProtocolMajor = TWON_PROTOCOLMAJOR
        .ProtocolMinor = TWON_PROTOCOLMINOR
        .SupportedGroups1 = DG_CONTROL Or DG_IMAGE
    End With

    sTmp = "LMik"
    Call CopyMemory(VarPtr(m_tAppID.Manufacturer(1)), StrPtr(StrConv(sTmp, vbFromUnicode)), Len(sTmp))
    sTmp = "VB Module"
    Call CopyMemory(VarPtr(m_tAppID.ProductFamily(1)), StrPtr(StrConv(sTmp, vbFromUnicode)), Len(sTmp))
    sTmp = "VB Module for TWAIN"
    Call CopyMemory(VarPtr(m_tAppID.ProductName(1)), StrPtr(StrConv(sTmp, vbFromUnicode)), Len(sTmp))

    '---------------------------------------
    ' Open Data Source Manager
    '   DG_CONTROL, DAT_PARENT, MSG_OPENDSM
    '---------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_PARENT, MSG_OPENDSM, m_lHndMsgWin)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace

    OpenTwainDSM = 0
    Exit Function

ErrPlace:
    OpenTwainDSM = 1
End Function

Private Function OpenTwainDS() As Long
    '--------------------------------------
    ' Function opens the Data Source
    ' Function returns 0 if OK, 1 if ERROR
    '--------------------------------------

    Dim iRtn As Integer

    On Local Error GoTo ErrPlace


    '-----------------------------------------------------------------------
    ' Open Data Source
    '   DG_CONTROL, DAT_IDENTITY, MSG_OPENDS
    '
    ' The default data source is opened. If you want user to select the new
    '   default one, call public function PopupSelectSourceDialog.
    '-----------------------------------------------------------------------
    Call ZeroMemory(VarPtr(m_tSrcID), Len(m_tSrcID))
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, MSG_OPENDS, m_tSrcID)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace

    OpenTwainDS = 0
    Exit Function

ErrPlace:
    OpenTwainDS = 1
End Function

Private Function CloseTwainDS() As Long
    '-----------------------------------------
    ' Function closes the Data Source Manager
    ' Function returns 0 if OK, 1 if ERROR
    '-----------------------------------------

    Dim iRtn As Integer

    On Local Error GoTo ErrPlace


    '-----------------------------------------
    ' Close Data Source
    '   DG_CONTROL, DAT_IDENTITY, MSG_CLOSEDS
    '-----------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, MSG_CLOSEDS, m_tSrcID)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace

    CloseTwainDS = 0
    Exit Function

ErrPlace:
    CloseTwainDS = 1
End Function

Private Function CloseTwainDSM() As Long
    '--------------------------------------
    ' Function closes the Data Source
    ' Function returns 0 if OK, 1 if ERROR
    '--------------------------------------

    Dim iRtn As Integer
    Dim lRtn As Long

    On Local Error GoTo ErrPlace


    '----------------------------------------
    ' Close Data Source Manager
    '   DG_CONTROL, DAT_PARENT, MSG_CLOSEDSM
    '----------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_PARENT, MSG_CLOSEDSM, m_lHndMsgWin)
    If iRtn <> TWRC_SUCCESS Then
        Call DestroyWindow(m_lHndMsgWin)
        GoTo ErrPlace
    End If

    '----------------------------
    ' Destroy the message window
    '----------------------------
    lRtn = DestroyWindow(m_lHndMsgWin)
    If lRtn = 0 Then GoTo ErrPlace

    CloseTwainDSM = 0
    Exit Function

ErrPlace:
    CloseTwainDSM = 1
End Function

Private Function SetImageSize(ByRef sngLeft As Single, _
                              ByRef sngTop As Single, _
                              ByRef sngRight As Single, _
                              ByRef sngBottom As Single) As Long
    '---------------------------------------------
    ' Function sets the size of the scanned image
    ' Function returns 0 if OK, 1 if ERROR
    '---------------------------------------------

    Dim tImageLayout As TW_IMAGELAYOUT
    Dim iRtn As Integer

    On Local Error GoTo ErrPlace


    '---------------------------------------------------------------------
    ' Set the size of the image - in default units
    '   DG_IMAGE, DAT_IMAGELAYOUT, MSG_SET
    '
    ' If you do not select any units, the INCHES are selected as default.
    ' Single type is converted into TWAIN TW_FIX32.
    '---------------------------------------------------------------------
    tImageLayout.Frame.Left = FloatToFix32(sngLeft)
    tImageLayout.Frame.Top = FloatToFix32(sngTop)
    tImageLayout.Frame.Right = FloatToFix32(sngRight)
    tImageLayout.Frame.Bottom = FloatToFix32(sngBottom)

    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_IMAGE, DAT_IMAGELAYOUT, MSG_SET, tImageLayout)
    If (iRtn <> TWRC_SUCCESS) And (iRtn <> TWRC_CHECKSTATUS) Then GoTo ErrPlace

    SetImageSize = 0
    Exit Function

ErrPlace:
    SetImageSize = 1
End Function

Private Function TwainTransfer(ByRef blShowUI As Boolean, _
                               ByRef sFileName As String, _
                               ByRef lCounter As Long) As Long

    Dim tUI As TW_USERINTERFACE
    Dim tMSG As MSG
    Dim tEvent As TW_EVENT
    Dim tPending As TW_PENDINGXFERS
    Dim lhDIB As Long
    Dim lRtn As Long
    Dim iRtn As Integer
    Dim sFileNameTmp As String

    On Local Error GoTo ErrPlace


    '----------------------------------------------
    ' Set tUI.ShowUI to 1 (show UI) or 0 (hide UI)
    '----------------------------------------------
    With tUI
        .ShowUI = IIf(blShowUI = True, 1, 0)
        .ModalUI = 1
        .hParent = m_lHndMsgWin
    End With

    '-----------------------------------------------
    ' Enable Data Source User Interface
    '   DG_CONTROL, DAT_USERINTERFACE, MSG_ENABLEDS
    '-----------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, MSG_ENABLEDS, tUI)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace


MSGLOOP:
    '------------------------------------------------------------------
    ' Process events in the message loop
    '   DG_CONTROL, DAT_EVENT, MSG_PROCESSEVENT
    '
    ' There are two messages we are interested in in this message loop
    '   - MSG_XFERREADY  - the data source is ready to transfer
    '   - MSG_CLOSEDSREQ - the data source requests to close itself
    '------------------------------------------------------------------
    While GetMessage(tMSG, 0&, 0&, 0&)
        Call ZeroMemory(VarPtr(tEvent), Len(tEvent))
        tEvent.pEvent = VarPtr(tMSG)
        iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_EVENT, MSG_PROCESSEVENT, tEvent)
        Select Case tEvent.TWMessage
            Case MSG_XFERREADY
                GoTo MSGGET
            Case MSG_CLOSEDSREQ
                GoTo MSGDISABLEDS
        End Select
        If iRtn = TWRC_NOTDSEVENT Then
            lRtn = TranslateMessage(tMSG)
            lRtn = DispatchMessage(tMSG)
        End If
    Wend


MSGGET:
    '-----------------------------------------------------------
    ' Start transfer
    '   DG_IMAGE, DAT_IMAGENATIVEXFER, MSG_GET
    '
    ' If transfer is successful, you will get the handle to DIB
    '-----------------------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_IMAGE, DAT_IMAGENATIVEXFER, MSG_GET, lhDIB)
    Select Case iRtn
        Case TWRC_XFERDONE
            '---------------------------------------------
            ' End transfer
            '   DG_CONTROL, DAT_PENDINGXFERS, MSG_ENDXFER
            '---------------------------------------------
            iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_PENDINGXFERS, MSG_ENDXFER, tPending)
            If iRtn <> TWRC_SUCCESS Then
                iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_PENDINGXFERS, MSG_RESET, tPending)
                iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, MSG_DISABLEDS, tUI)
                GoTo ErrPlace
            End If

            '-------------------------------
            ' Save DIB handle into BMP file
            '-------------------------------
            sFileNameTmp = sFileName & "_" & Format$(lCounter, "000000") & ".bmp"
            lRtn = SaveDIBToFile(lhDIB, sFileNameTmp)
            If lRtn Then
                iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_PENDINGXFERS, MSG_RESET, tPending)
                iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, MSG_DISABLEDS, tUI)
                GoTo ErrPlace
            End If
            '----------------------------------------------------
            ' Increment Filename counter after successful saving
            '----------------------------------------------------
            lCounter = lCounter + 1
        Case TWRC_CANCEL
            '---------------------------------------------
            ' End transfer
            '   DG_CONTROL, DAT_PENDINGXFERS, MSG_ENDXFER
            '---------------------------------------------
            iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_PENDINGXFERS, MSG_ENDXFER, tPending)
            If iRtn <> TWRC_SUCCESS Then
                iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_PENDINGXFERS, MSG_RESET, tPending)
                iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, MSG_DISABLEDS, tUI)
                GoTo ErrPlace
            End If
            '-----------------
            ' Free DIB handle
            '-----------------
            lRtn = GlobalFree(lhDIB)
        Case Default    'TWRC_FAILURE and other scenarios
            iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_PENDINGXFERS, MSG_RESET, tPending)
            iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, MSG_DISABLEDS, tUI)
            GoTo ErrPlace
    End Select

    '-------------------------------------------
    ' Continue scanning if any pending document
    '-------------------------------------------
    If tPending.Count <> 0 Then GoTo MSGGET
    'If blShowUI Then GoTo MSGLOOP


MSGDISABLEDS:
    '------------------------------------------------
    ' Disable Data Source
    '   DG_CONTROL, DAT_USERINTERFACE, MSG_DISABLEDS
    '------------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, MSG_DISABLEDS, tUI)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace

    TwainTransfer = 0
    Exit Function

ErrPlace:
    If lhDIB Then lRtn = GlobalFree(lhDIB)
    TwainTransfer = 1
End Function

Private Function SaveDIBToFile(ByRef lhDIB As Long, _
                               ByRef sFileName As String) As Long
    '-------------------------------------------------------------------------
    ' Function saves the DIB handle (device independent bitmap) into BMP file
    ' Function returns 0 if OK, 1 if ERROR
    '-------------------------------------------------------------------------

    Dim lpDIB As Long
    Dim tBIH As BITMAPINFOHEADER
    Dim lDIBOffset As Long
    Dim tRGB As RGBQUAD
    Dim lDIBWidth As Long
    Dim lDIBSize As Long
    Dim bDIBits() As Byte
    Dim lRtn As Long
    Dim tBFH As BITMAPFILEHEADER
    Dim iFileNum As Integer

    On Local Error GoTo ErrPlace


    If sFileName = "" Then GoTo ErrPlace

    If Dir(sFileName, vbNormal Or vbHidden Or vbSystem) <> "" Then
        Call SetAttr(sFileName, vbNormal)
        Call Kill(sFileName)
    End If

    lpDIB = GlobalLock(lhDIB)
    If lpDIB = 0 Then GoTo ErrPlace

    Call CopyMemory(VarPtr(tBIH), lpDIB, Len(tBIH))
    If (tBIH.biBitCount = 1) Or (tBIH.biBitCount = 8) Then tBIH.biClrUsed = 2 ^ tBIH.biBitCount
    lDIBOffset = Len(tBIH) + (tBIH.biClrUsed * Len(tRGB))
    lDIBWidth = (((tBIH.biWidth * tBIH.biBitCount) + 31) \ 32) * 4
    lDIBSize = lDIBOffset + (lDIBWidth * tBIH.biHeight)

    ReDim bDIBits(1 To lDIBSize) As Byte
    Call CopyMemory(VarPtr(bDIBits(1)), lpDIB, lDIBSize)

    lRtn = GlobalUnlock(lhDIB)
    lRtn = GlobalFree(lhDIB)
    lhDIB = 0

    With tBFH
        .bfType = 19778  ' "BM"
        .bfSize = Len(tBFH) + lDIBSize
        .bfOffBits = Len(tBFH) + lDIBOffset
    End With

    iFileNum = FreeFile
    Open sFileName For Binary As #iFileNum
        Put #iFileNum, , tBFH
        Put #iFileNum, , bDIBits()
    Close #iFileNum

    SaveDIBToFile = 0
    Exit Function

ErrPlace:
    lRtn = GlobalUnlock(lhDIB)
    lRtn = GlobalFree(lhDIB)
    lhDIB = 0
    SaveDIBToFile = 1
End Function

Private Function TwainGetOnevalue(ByVal Cap As TWAIN_MDL_CAPABILITY, _
                                  ByRef Item As Variant) As Long
    '------------------------------------------------------------------------
    ' Function gets value of ONEVALUE type capability
    '
    ' There are four types of containers that TWAIN defines for capabilities
    '   (TW_ONEVALUE, TW_ARRAY, TW_RANGE and TW_ENUMERATION)
    ' This module deals with one of them only - TW_ONEVALUE (single value)
    ' To get some capability you have to fill TW_ONEVALUE fields and use
    '   the triplet DG_CONTROL DAT_CAPABILITY MSG_GET
    ' The macros that convert some data types are used here as well
    '
    ' Function returns 0 if OK, 1 if an error occurs
    '------------------------------------------------------------------------

    Dim tCapability As TW_CAPABILITY
    Dim iRtn As Integer
    Dim lpOneValue As Long
    Dim tOneValue As TW_ONEVALUE
    Dim tFix32 As TW_FIX32

    Dim lRtn As Long

    On Local Error GoTo ErrPlace


    tCapability.Cap = Cap
    tCapability.ConType = TWON_ONEVALUE

    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_CAPABILITY, MSG_GET, tCapability)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace

    lpOneValue = GlobalLock(tCapability.hContainer)
    Call CopyMemory(VarPtr(tOneValue.ItemType), lpOneValue, 2&)  ' TW_UINT16
    Call CopyMemory(VarPtr(tOneValue.Item), lpOneValue + 2, 4&)  ' TW_UINT32
    lRtn = GlobalUnlock(tCapability.hContainer)
    lRtn = GlobalFree(tCapability.hContainer)

    Select Case tOneValue.ItemType
        Case INT16
            Item = tOneValue.Item
        Case UINT16, BOOL
            Item = FromUnsignedShort(CInt(tOneValue.Item))
        Case FIX32
            Call CopyMemory(VarPtr(tFix32), VarPtr(tOneValue.Item), 4&)
            Item = Fix32ToFloat(tFix32)
    End Select

    TwainGetOnevalue = 0
    Exit Function

ErrPlace:
    TwainGetOnevalue = 1
End Function

Private Function TwainSetOnevalue(ByVal Cap As TWAIN_MDL_CAPABILITY, _
                                  ByVal ItemType As TWAIN_MDL_ITEMYPE, _
                                  ByRef Item As Variant) As Long
    '-----------------------------------------------------------------------
    ' Function sets value of ONEVALUE type capability
    '
    ' There are four types of container that TWAIN defines for capabilities
    '   (TW_ONEVALUE, TW_ARRAY, TW_RANGE and TW_ENUMERATION)
    ' This module deals with one of them only - TW_ONEVALUE (single value)
    ' To set a capability you have to fill TW_ONEVALUE fields and use
    '   the triplet DG_CONTROL DAT_CAPABILITY MSG_SET
    ' The macros that convert some data types are used here as well
    '
    ' Function returns 0 if OK, 1 if an error occurs
    '-----------------------------------------------------------------------

    Dim tCapability As TW_CAPABILITY
    Dim tOneValue As TW_ONEVALUE
    Dim iTmp As Integer
    Dim tFix32 As TW_FIX32
    Dim lhOneValue As Long
    Dim lpOneValue As Long
    Dim lRtn As Long
    Dim iRtn As Integer

    On Local Error GoTo ErrPlace


    tCapability.Cap = Cap
    tCapability.ConType = TWON_ONEVALUE

    tOneValue.ItemType = ItemType

    Select Case ItemType
        Case INT16
            tOneValue.Item = CInt(Item)
        Case UINT16, BOOL
            iTmp = ToUnsignedShort(CLng(Item))
            Call CopyMemory(VarPtr(tOneValue.Item), VarPtr(iTmp), 2&)
        Case FIX32
            tFix32 = FloatToFix32(CSng(Item))
            Call CopyMemory(VarPtr(tOneValue.Item), VarPtr(tFix32), 4&)
    End Select

    lhOneValue = GlobalAlloc(GHND, Len(tOneValue))
    lpOneValue = GlobalLock(lhOneValue)
    Call CopyMemory(lpOneValue, VarPtr(tOneValue.ItemType), 2&)  ' TW_UINT16
    Call CopyMemory(lpOneValue + 2, VarPtr(tOneValue.Item), 4&)  ' TW_UINT32
    lRtn = GlobalUnlock(lhOneValue)
    tCapability.hContainer = lhOneValue

    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_CAPABILITY, MSG_SET, tCapability)
    If iRtn <> TWRC_SUCCESS Then
        lRtn = GlobalFree(lhOneValue)
        GoTo ErrPlace
    End If
    lRtn = GlobalFree(lhOneValue)

    TwainSetOnevalue = 0
    Exit Function

ErrPlace:
    TwainSetOnevalue = 1
End Function

Private Function ToUnsignedShort(ByRef lSrc As Long) As Integer
    '-------------------------------------------------------------------------
    ' Function puts number ranging from 0 to 65535 into 2-byte VB Integer
    ' (useful for communicating with other dll that uses unsigned data types)
    '
    ' Function returns unsigned 2-byte value in VB Integer type
    '-------------------------------------------------------------------------

    Dim iTmp As Integer


    If (lSrc < 0) Or (lSrc > 65535) Then
        iTmp = 0
    Else
        Call CopyMemory(VarPtr(iTmp), VarPtr(lSrc), 2&)
    End If

    ' Another way
    'iTmp = IIf(lSrc > 32767, lSrc - 65536, lSrc)

    ToUnsignedShort = iTmp
End Function

Private Function FromUnsignedShort(ByRef iSrc As Integer) As Long
    '-------------------------------------------------------------------------
    ' Function gets the 2-byte unsigned number from VB Integer data type
    ' (useful for communicating with other dll that uses unsigned data types)
    '
    ' Function returns unsigned 2-byte value in VB Long type
    '-------------------------------------------------------------------------

    Dim lTmp As Long


    Call CopyMemory(VarPtr(lTmp), VarPtr(iSrc), 2&)

    ' Another way
    'lTmp = IIf(iSrc < 0, iSrc + 65536, iSrc)

    FromUnsignedShort = lTmp
End Function

Private Function ToUnsignedLong(ByRef vSrc As Variant) As Long
    '-------------------------------------------------------------------------
    ' Function puts number ranging from 0 to 4294967295 into 4-byte VB Long
    ' (useful for communicating with other dll that uses unsigned data types)
    '
    ' Function returns unsigned 4-byte value in VB Long type
    '-------------------------------------------------------------------------

    Dim lTmp As Long


    If (vSrc < 0) Or (vSrc > 4294967295#) Then
        lTmp = 0
    Else
        lTmp = IIf(vSrc > 2147483647, vSrc - 4294967296#, vSrc)
    End If

    ToUnsignedLong = lTmp
End Function

Private Function FromUnsignedLong(ByRef lSrc As Long) As Variant
    '-------------------------------------------------------------------------
    ' Function gets the 4-byte unsigned number from VB Long data type
    ' (useful for communicating with other dll that uses unsigned data types)
    '
    ' Function returns unsigned 4-byte value in VB Variant type
    '-------------------------------------------------------------------------

    Dim vTmp As Variant


    vTmp = IIf(lSrc < 0, lSrc + 4294967296#, lSrc)

    FromUnsignedLong = vTmp
End Function

Private Function Fix32ToFloat(ByRef tFix32 As TW_FIX32) As Single
    '--------------------------------------------------------------------------
    ' Function converts TWAIN TW_FIX32 data structure into VB Single data type
    ' (needed for communicating with TWAIN)
    '
    ' Function returns floating-point number in VB Single data type
    '--------------------------------------------------------------------------

    Dim sngTmp As Single


    sngTmp = tFix32.Whole + CSng(FromUnsignedShort(tFix32.Frac) / 65536)

    Fix32ToFloat = sngTmp
End Function

Private Function FloatToFix32(ByRef sngSrc As Single) As TW_FIX32
    '--------------------------------------------------------------------------
    ' Function converts VB Single data type into TWAIN TW_FIX32 data structure
    ' (needed for communicating with TWAIN)
    '
    ' Function returns TW_FIX32 data structure
    '--------------------------------------------------------------------------

    Dim tFix32 As TW_FIX32


    tFix32.Whole = CInt(Fix(sngSrc))
    tFix32.Frac = ToUnsignedShort(CLng(sngSrc * 65536) And 65535)

    FloatToFix32 = tFix32
End Function
