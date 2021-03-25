Attribute VB_Name = "mdlTwain"
'*******************************************************************************
'
' Description: VB Module for accessing TWAIN compatible scanner (VB 5, 6)
'
' Author:      Lumir Mik (lmik@seznam.cz)
'
' Version:     1.0
'
' License:     Free to any use. If you change some part of this code, please,
'              mention it here.
'              Receive it as my contribution to free programmer sources
'              in which I found much help and inspiration.
'
' There are 3 public functions in this module:
'
'   1. PopupSelectSourceDialog
'           shows TWAIN dialog for selecting default source for acquisition
'
'   2. TransferWithoutUI
'           transfers one image from TWAIN data source without showing
'           the data source user interface (silent transfer). The programmer
'           can set following attributes of the image:
'               - resolution (DPI)
'               - colour depth - monochromatic, grey, fullcolour
'               - image size and position on the scanner glass
'                       - left, top, right, bottom (in inches).
'           The image is saved into the BMP file.
'
'   3. TransferWithUI
'           transfers one image from TWAIN data source using the data
'           source user interface to set image attributes.
'           The image is saved into the BMP file.
'
'*******************************************************************************

Option Explicit

'-----------------------------
' Declaration for TWAIN_32.DLL
'-----------------------------
Private Declare Function DSM_Entry Lib "Twain_32.dll" _
                                   (ByRef pOrigin As Any, _
                                    ByRef pDest As Any, _
                                    ByVal DG As Long, _
                                    ByVal DAT As Integer, _
                                    ByVal MSG As Integer, _
                                    ByRef pData As Any) As Integer

Private Type TW_VERSION
    MajorNum        As Integer                  ' TW_UINT16
    MinorNum        As Integer                  ' TW_UINT16
    Language        As Integer                  ' TW_UINT16
    Country         As Integer                  ' TW_UINT16
    Info(1 To 34)   As Byte                     ' TW_STR32
End Type

Private Type TW_IDENTITY
    Id                      As Long             ' TW_UINT32
    Version                 As TW_VERSION       ' TW_VERSION
    ProtocolMajor           As Integer          ' TW_UINT16
    ProtocolMinor           As Integer          ' TW_UINT16
    SupportedGroups1        As Integer          ' TW_UINT32
    SupportedGroups2        As Integer
    Manufacturer(1 To 34)   As Byte             ' TW_STR32
    ProductFamily(1 To 34)  As Byte             ' TW_STR32
    ProductName(1 To 34)    As Byte             ' TW_STR32
End Type

Private Type TW_USERINTERFACE
    ShowUI   As Integer                         ' TW_BOOL
    ModalUI  As Integer                         ' TW_BOOL
    hParent  As Long                            ' TW_HANDLE
End Type

Private Type TW_PENDINGXFERS
    Count       As Integer                      ' TW_UINT16
    Reserved1   As Integer                      ' TW_UINT32
    Reserved2   As Integer
End Type

Private Type TW_ONEVALUE
    ItemType As Integer                         ' TW_UINT16
    Item1    As Integer                         ' TW_UINT32
    Item2    As Integer
End Type

Private Type TW_CAPABILITY
    Cap          As Integer                     ' TW_UINT16
    ConType      As Integer                     ' TW_UINT16
    hContainer   As Long                        ' TW_HANDLE
End Type

Private Type TW_FIX32
    Whole   As Integer                          ' TW_INT16
    Frac    As Integer                          ' TW_UINT16
End Type

Private Type TW_FRAME
    Left     As TW_FIX32                        ' TW_FIX32
    Top      As TW_FIX32                        ' TW_FIX32
    Right    As TW_FIX32                        ' TW_FIX32
    Bottom   As TW_FIX32                        ' TW_FIX32
End Type

Private Type TW_IMAGELAYOUT
    Frame            As TW_FRAME                ' TW_FRAME
    DocumentNumber   As Long                    ' TW_UINT32
    PageNumber       As Long                    ' TW_UINT32
    FrameNumber      As Long                    ' TW_UINT32
End Type

Private Type TW_EVENT
    pEvent      As Long                         ' TW_MEMREF
    TWMessage   As Integer                      ' TW_UINT16
End Type

Private Const DG_CONTROL = 1
Private Const DG_IMAGE = 2

Private Const MSG_GET = 1
Private Const MSG_SET = 6
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
Private Const TWRC_CHECKSTATUS = 2
Private Const TWRC_DSEVENT = 4
Private Const TWRC_NOTDSEVENT = 5
Private Const TWRC_XFERDONE = 6

Private Const TWLG_CZECH = 45

Private Const TWCY_CZECHOSLOVAKIA = 42

Private Const TWON_PROTOCOLMAJOR = 1
Private Const TWON_ONEVALUE = 5
Private Const TWON_PROTOCOLMINOR = 9


'-------------------------
' Declaration for WIN32API
'-------------------------
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
                               (ByVal pDest As Long, _
                                ByVal pSource As Long, _
                                ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" _
                               (ByVal pDest As Long, _
                                ByVal Length As Long)
Private Declare Function GlobalFree Lib "kernel32.dll" _
                                    (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" _
                                    (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" _
                                      (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" _
                                     (ByVal wFlags As Long, _
                                      ByVal dwBytes As Long) As Long
Private Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" _
                                    (ByRef lpMsg As MSG, _
                                     ByVal hwnd As Long, _
                                     ByVal wMsgFilterMin As Long, _
                                     ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32.dll" _
                                          (ByRef lpMsg As MSG) As Long
Private Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" _
                                         (ByRef lpMsg As MSG) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" _
                                        (ByVal dwExStyle As Long, _
                                         ByVal lpClassName As String, _
                                         ByVal lpWindowName As String, _
                                         ByVal dwStyle As Long, _
                                         ByVal x As Long, _
                                         ByVal y As Long, _
                                         ByVal nWidth As Long, _
                                         ByVal nHeight As Long, _
                                         ByVal hWndParent As Long, _
                                         ByVal hMenu As Long, _
                                         ByVal hInstance As Long, _
                                         ByVal lpParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32.dll" _
                                       (ByVal hwnd As Long) As Long

Private Type BITMAPFILEHEADER
    bfType      As Integer
    bfSize      As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits   As Long
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type RGBQUAD
    rgbBlue     As Byte
    rgbGreen    As Byte
    rgbRed      As Byte
    rgbReserved As Byte
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MSG
    hwnd    As Long
    message As Long
    wParam  As Long
    lParam  As Long
    time    As Long
    pt      As POINTAPI
End Type

Private Const GHND = 66


'----------------------------
' Declaration for this Module
'----------------------------
Private m_tAppID As TW_IDENTITY
Private m_tSrcID As TW_IDENTITY
Private m_lHndMsgWin As Long

Public Enum TWAIN_MDL_COLOURTYPE
    BW = 0                              ' TWPT_BW
    GREY = 1                            ' TWPT_GRAY
    RGB = 2                             ' TWPT_RGB
End Enum

Private Enum TWAIN_MDL_CAPABILITY
    XFERCOUNT = 1                       ' CAP_XFERCOUNT
    PIXELTYPE = 257                     ' ICAP_PIXELTYPE
    INDICATORS = 4107                   ' CAP_INDICATORS
    UICONTROLLABLE = 4110               ' CAP_UICONTROLLABLE
    PHYSICALWIDTH = 4369                ' ICAP_PSYSICALWIDTH
    PHYSICALHEIGHT = 4370               ' ICAP_PSYSICALHEIGHT
    XRESOLUTION = 4376                  ' ICAP_XRESOLUTION
    YRESOLUTION = 4377                  ' ICAP_YRESOLUTION
    BITDEPTH = 4395                     ' ICAP_BITDEPTH
End Enum

Private Enum TWAIN_MDL_ITEMYPE
    INT16 = 1                           ' TW_INT16      short
    UINT16 = 4                          ' TW_UINT16     unsigned short
    BOOL = 6                            ' TW_BOOL       unsigned short
    FIX32 = 7                           ' TW_FIX32      structure
End Enum

Public Function TransferWithoutUI(ByVal sngResolution As Single, _
                                  ByVal tColourType As TWAIN_MDL_COLOURTYPE, _
                                  ByVal sngImageLeft As Single, _
                                  ByVal sngImageTop As Single, _
                                  ByVal sngImageRight As Single, _
                                  ByVal sngImageBottom As Single, _
                                  ByVal sBMPFileName As String) As Long

    '----------------------------------------------------------------------------
    ' Function transfers one image from Twain data source without showing
    '   the data source user interface (silent transfer).
    '
    ' Input values
    '   - sngResolution (Single) - resolution of the image in DPI
    '                              (dots per inch)
    '   - tColourType (UDT) - colour depth of the imaged - monochromatic (BW),
    '                         colours of grey (GREY), full colours (COLOUR)
    '   - sngImageLeft, sngImageTop, sngImageRight, sngImageBottom (Single) -
    '       values determine the rectangle on the scanner glass that will
    '       be scanned (default units are inches) - if you set Right and Bottom
    '       values to 0, the module sets maximum values the scanner driver allows
    '       (the bottom right corner of the scanner glass)
    '   - sBMPFileName (String) - the file name of the saved image
    '
    ' Function returns 0 if OK, 1 if an error occurs
    '----------------------------------------------------------------------------
                                   
    Dim lRtn As Long
    Dim lTmp As Long
    Dim blTwainOpen As Boolean
    Dim lhDIB As Long
    
    On Local Error GoTo ErrPlace

    '-------------------------------
    ' Open Twain Data Source Manager
    '-------------------------------
    lRtn = OpenTwainDSM()
    If lRtn Then GoTo ErrPlace
    blTwainOpen = True

    '-----------------------
    ' Open Twain Data Source
    '-----------------------
    lRtn = OpenTwainDS()
    If lRtn Then GoTo ErrPlace

    '-----------------------------------------------------------
    ' Set all important attributes of the image and the transfer
    '-----------------------------------------------------------
    
    '----------------------------------------------------------------------
    ' Set image size and position
    ' If sngImageRight or sngImageBottom is 0 put physical width and height
    '   of the scanner into these values
    '----------------------------------------------------------------------
    If (sngImageRight = 0) Or (sngImageBottom = 0) Then
        lRtn = TwainGetOneValue(PHYSICALWIDTH, sngImageRight)
        If lRtn Then GoTo ErrPlace
        lRtn = TwainGetOneValue(PHYSICALHEIGHT, sngImageBottom)
        If lRtn Then GoTo ErrPlace
    End If

    lRtn = SetImageSize(sngImageLeft, sngImageTop, sngImageRight, sngImageBottom)
    If lRtn Then GoTo ErrPlace

    '-----------------------------------------------
    ' Set the image resolution in DPI - both X and Y
    '-----------------------------------------------
    lRtn = TwainSetOneValue(XRESOLUTION, FIX32, sngResolution)
    If lRtn Then GoTo ErrPlace

    lRtn = TwainSetOneValue(YRESOLUTION, FIX32, sngResolution)
    If lRtn Then GoTo ErrPlace

    '--------------------------
    ' Set the image colour type
    '--------------------------
    lRtn = TwainSetOneValue(PIXELTYPE, UINT16, tColourType)
    If lRtn Then GoTo ErrPlace
    
    '----------------------------------------------------------------
    ' If the colour type is fullcolour, set the bitdepth of the image
    '   - 24 bits, 32 bits, ...
    '----------------------------------------------------------------
    If tColourType = RGB Then lRtn = TwainSetOneValue(BITDEPTH, UINT16, 24)

    '---------------------------------------------------
    ' Set number of images you want to transfer (just 1)
    '---------------------------------------------------
    lRtn = TwainSetOneValue(XFERCOUNT, INT16, 1)
    If lRtn Then GoTo ErrPlace

    '----------------------------------------------------
    ' TRANSFER the image with UI disabled.
    '   If successful, lhDIB is filled with handle to DIB
    '----------------------------------------------------
    lRtn = TwainTransfer(False, lhDIB)
    If lRtn Then GoTo ErrPlace

    '------------------
    ' Close Data Source
    '------------------
    lRtn = CloseTwainDS()
    If lRtn Then GoTo ErrPlace
    
    '--------------------------
    ' Close Data Source Manager
    '--------------------------
    lRtn = CloseTwainDSM()
    If lRtn Then GoTo ErrPlace
    blTwainOpen = False
    
    '----------------------------------
    ' Save DIB handle into the BMP file
    '----------------------------------
    lRtn = SaveDIBToFile(lhDIB, sBMPFileName)
    If lRtn Then GoTo ErrPlace

    TransferWithoutUI = 0
    Exit Function

ErrPlace:
    If lhDIB Then lRtn = GlobalFree(lhDIB)
    If blTwainOpen Then lRtn = CloseTwainDS(): lRtn = CloseTwainDSM()
    TransferWithoutUI = 1
End Function

Public Function TransferWithUI(ByVal sBMPFileName As String) As Long
        
    '-------------------------------------------------------------------
    ' Function transfers one image from Twain data source using the data
    '   source user interface to set image attributes.
    '
    ' Input values
    '   - sBMPFileName (String) - the file name of the saved image
    '
    ' Function returns 0 if OK, 1 if an error occurs
    '-------------------------------------------------------------
    
    Dim lRtn As Long
    Dim blTwainOpen As Boolean
    Dim lhDIB As Long
    
    On Local Error GoTo ErrPlace

    '-------------------------------
    ' Open Twain Data Source Manager
    '-------------------------------
    lRtn = OpenTwainDSM()
    If lRtn Then GoTo ErrPlace
    blTwainOpen = True

    '-----------------------
    ' Open Twain Data Source
    '-----------------------
    lRtn = OpenTwainDS()
    If lRtn Then GoTo ErrPlace

    '----------------------------------------------------
    ' TRANSFER the image with UI enabled.
    '   If successful, lhDIB is filled with handle to DIB
    '----------------------------------------------------
    lRtn = TwainTransfer(True, lhDIB)
    If lRtn Then GoTo ErrPlace

    '------------------
    ' Close Data Source
    '------------------
    lRtn = CloseTwainDS()
    If lRtn Then GoTo ErrPlace
    
    '--------------------------
    ' Close Data Source Manager
    '--------------------------
    lRtn = CloseTwainDSM()
    If lRtn Then GoTo ErrPlace
    blTwainOpen = False
    
    '----------------------------------
    ' Save DIB handle into the BMP file
    '----------------------------------
    lRtn = SaveDIBToFile(lhDIB, sBMPFileName)
    If lRtn Then GoTo ErrPlace

    TransferWithUI = 0
    Exit Function

ErrPlace:
    If lhDIB Then lRtn = GlobalFree(lhDIB)
    If blTwainOpen Then lRtn = CloseTwainDS(): lRtn = CloseTwainDSM()
    TransferWithUI = 1
End Function

Public Function PopupSelectSourceDialog() As Long
    
    '------------------------------------------------------------------
    ' Function shows the Twain dialog for selecting default data source
    '
    ' Function returns 0 if OK, 1 if an error occurs
    '------------------------------------------------------------------
    
    Dim iRtn As Integer
    Dim lRtn As Long
    
    On Local Error GoTo ErrPlace
    
    '-------------------------------
    ' Open Twain Data Source Manager
    '-------------------------------
    lRtn = OpenTwainDSM()
    If lRtn Then GoTo ErrPlace

    '----------------------------------------------------
    ' Popup "Select source" dialog
    '   DG_CONTROL, DAT_IDENTITY, MSG_USERSELECT
    '----------------------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, _
                     MSG_USERSELECT, m_tSrcID)
    If iRtn <> TWRC_SUCCESS Then
        lRtn = CloseTwainDSM()
        GoTo ErrPlace
    End If
    
    '--------------------------------
    ' Close Twain Data Source Manager
    '--------------------------------
    lRtn = CloseTwainDSM()
    If lRtn Then GoTo ErrPlace
    
    PopupSelectSourceDialog = 0
    Exit Function
    
ErrPlace:
    PopupSelectSourceDialog = 1
End Function

Private Function OpenTwainDSM() As Long
    
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    '----------------------------------------------------
    ' Create window that will receive all TWAIN messages
    ' Message loop can be found in TwainTransfer function
    '----------------------------------------------------
    m_lHndMsgWin = CreateWindowEx(0&, "#32770", "TWAIN_MSG_WINDOW", 0&, _
                                  10&, 10&, 150&, 50&, 0&, 0&, 0&, 0&)
    If m_lHndMsgWin = 0 Then GoTo ErrPlace
    
    '------------------------------------------------------------
    ' Introduce yourself to TWAIN - MajorNum, MinorNum, Language,
    ' Country, Manufacturer, ProductFamily, ProductName, etc.
    '------------------------------------------------------------
    Call ZeroMemory(VarPtr(m_tAppID), Len(m_tAppID))
    With m_tAppID
        .Version.MajorNum = 1
        .Version.Language = TWLG_CZECH
        .Version.Country = TWCY_CZECHOSLOVAKIA
        .ProtocolMajor = TWON_PROTOCOLMAJOR
        .ProtocolMinor = TWON_PROTOCOLMINOR
        .SupportedGroups1 = DG_CONTROL Or DG_IMAGE
    End With
    
    Call CopyMemory(VarPtr(m_tAppID.Manufacturer(1)), _
                    StrPtr(StrConv("LMik", vbFromUnicode)), _
                    Len("LMik"))
    Call CopyMemory(VarPtr(m_tAppID.ProductFamily(1)), _
                    StrPtr(StrConv("VB Module", vbFromUnicode)), _
                    Len("VB Module"))
    Call CopyMemory(VarPtr(m_tAppID.ProductName(1)), _
                    StrPtr(StrConv("VB Module for TWAIN", vbFromUnicode)), _
                    Len("VB Module for TWAIN"))
    
    '--------------------------------------
    ' Open Data Source Manager
    '   DG_CONTROL, DAT_PARENT, MSG_OPENDSM
    '--------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_PARENT, MSG_OPENDSM, _
                     m_lHndMsgWin)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace
    
    OpenTwainDSM = 0
    Exit Function
    
ErrPlace:
    OpenTwainDSM = 1
End Function

Private Function OpenTwainDS() As Long

    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    '----------------------------------------------------------------------
    ' Open Data Source
    '   DG_CONTROL, DAT_IDENTITY, MSG_OPENDS
    '
    ' The default data source is opened. If you want user to select the new
    '   default one, call public function PopupSelectSourceDialog.
    '----------------------------------------------------------------------
    Call ZeroMemory(VarPtr(m_tSrcID), Len(m_tSrcID))
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, MSG_OPENDS, _
                     m_tSrcID)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace

    OpenTwainDS = 0
    Exit Function
    
ErrPlace:
    OpenTwainDS = 1
End Function

Private Function CloseTwainDS() As Long
    
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    '----------------------------------------
    ' Close Data Source
    '   DG_CONTROL, DAT_IDENTITY, MSG_CLOSEDS
    '----------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, _
                     MSG_CLOSEDS, m_tSrcID)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace
    
    CloseTwainDS = 0
    Exit Function

ErrPlace:
    CloseTwainDS = 1
End Function

Private Function CloseTwainDSM() As Long
    
    Dim lRtn As Long
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace

    '---------------------------------------
    ' Close Data Source Manager
    '   DG_CONTROL, DAT_PARENT, MSG_CLOSEDSM
    '---------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_PARENT, MSG_CLOSEDSM, _
                     m_lHndMsgWin)
    If iRtn <> TWRC_SUCCESS Then
        lRtn = DestroyWindow(m_lHndMsgWin)
        GoTo ErrPlace
    End If

    '---------------------------
    ' Destroy the message window
    '---------------------------
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

    Dim tImageLayout As TW_IMAGELAYOUT
    Dim lRtn As Long
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    '-------------------------------------------------------------------
    ' Set the size of the image - in default units
    '   DG_IMAGE, DAT_IMAGELAYOUT, MSG_SET
    '
    ' If you do not select any units the INCHES are selected as default.
    ' The values of Single type are converted into TWAIN TW_FIX32.
    '-------------------------------------------------------------------
    lRtn = FloatToFix32(sngLeft, tImageLayout.Frame.Left)
    If lRtn Then GoTo ErrPlace
    
    lRtn = FloatToFix32(sngTop, tImageLayout.Frame.Top)
    If lRtn Then GoTo ErrPlace
    
    lRtn = FloatToFix32(sngRight, tImageLayout.Frame.Right)
    If lRtn Then GoTo ErrPlace
    
    lRtn = FloatToFix32(sngBottom, tImageLayout.Frame.Bottom)
    If lRtn Then GoTo ErrPlace
    
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_IMAGE, DAT_IMAGELAYOUT, MSG_SET, _
                     tImageLayout)
    If (iRtn <> TWRC_SUCCESS) And (iRtn <> TWRC_CHECKSTATUS) Then GoTo ErrPlace
    
    SetImageSize = 0
    Exit Function
    
ErrPlace:
    SetImageSize = 1
End Function

Private Function TwainTransfer(ByRef blShowUI As Boolean, _
                               ByRef lDIBHandle As Long) As Long
                               
    Dim tUI As TW_USERINTERFACE
    Dim tPending As TW_PENDINGXFERS
    Dim lhDIB As Long
    Dim tEvent As TW_EVENT
    Dim tMSG As MSG
    Dim lRtn As Long
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    '---------------------------------------------
    ' Set tUI.ShowUI to 1 (show UI) or 0 (hide UI)
    '---------------------------------------------
    With tUI
        .ShowUI = IIf(blShowUI = True, 1, 0)
        .ModalUI = 1
        .hParent = m_lHndMsgWin
    End With
    
    '----------------------------------------------
    ' Enable Data Source User Interface
    '   DG_CONTROL, DAT_USERINTERFACE, MSG_ENABLEDS
    '----------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, _
                     MSG_ENABLEDS, tUI)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace
    
    '-----------------------------------------------------------------
    ' Process events in the message loop
    '   DG_CONTROL, DAT_EVENT, MSG_PROCESSEVENT
    '
    ' There are two messages we are interested in in this message loop
    '   - MSG_XFERREADY - the data source is ready to transfer
    '   - MSG_CLOSEDSREQ - the data source requests to close itself
    '-----------------------------------------------------------------
    While GetMessage(tMSG, 0&, 0&, 0&)
        Call ZeroMemory(VarPtr(tEvent), Len(tEvent))
        tEvent.pEvent = VarPtr(tMSG)
        iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_EVENT, _
                         MSG_PROCESSEVENT, tEvent)
        Select Case tEvent.TWMessage
            Case MSG_XFERREADY
                GoTo MSGGET
            Case MSG_CLOSEDSREQ
                GoTo MSGDISABLEDS
        End Select
        lRtn = TranslateMessage(tMSG)
        lRtn = DispatchMessage(tMSG)
    Wend
    
MSGGET:
    '----------------------------------------------------
    ' Start transfer
    '   DG_IMAGE, DAT_IMAGENATIVEXFER, MSG_GET
    '
    ' If transfer is successful you get the handle to DIB
    '----------------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_IMAGE, DAT_IMAGENATIVEXFER, _
                     MSG_GET, lhDIB)
    If iRtn <> TWRC_XFERDONE Then
        iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_PENDINGXFERS, _
                         MSG_ENDXFER, tPending)
        iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, _
                         MSG_DISABLEDS, tUI)
        GoTo ErrPlace
    End If
    
    '--------------------------------------------
    ' End transfer
    '   DG_CONTROL, DAT_PENDINGXFERS, MSG_ENDXFER
    '--------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_PENDINGXFERS, _
                     MSG_ENDXFER, tPending)
    If iRtn <> TWRC_SUCCESS Then
        iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, _
                         MSG_DISABLEDS, tUI)
        GoTo ErrPlace
    End If
    
MSGDISABLEDS:
    '-----------------------------------------------
    ' Disable Data Source
    '   DG_CONTROL, DAT_USERINTERFACE, MSG_DISABLEDS
    '-----------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, _
                     MSG_DISABLEDS, tUI)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace

    lDIBHandle = lhDIB
    TwainTransfer = 0
    Exit Function
    
ErrPlace:
    If lhDIB Then lRtn = GlobalFree(lhDIB)
    lDIBHandle = 0
    TwainTransfer = 1
End Function

Private Function SaveDIBToFile(ByRef lhDIB As Long, _
                               ByRef sFileName As String) As Long
    
    '---------------------------------------------------------------------------
    ' Function saves the handle to DIB (device independent bitmap) into BMP file
    '---------------------------------------------------------------------------

    Dim tBFH As BITMAPFILEHEADER
    Dim tBIH As BITMAPINFOHEADER
    Dim tRGB As RGBQUAD
    Dim lpDIB As Long
    Dim lDIBSize As Long
    Dim bDIBits() As Byte
    Dim iFileNum As Integer
    Dim lRtn As Long
    
    On Local Error GoTo ErrPlace
    
    If sFileName = "" Then GoTo ErrPlace
    
    If Dir(sFileName, vbNormal Or vbHidden Or vbSystem) <> "" Then
        Call SetAttr(sFileName, vbNormal)
        Call Kill(sFileName)
    End If
    
    lpDIB = GlobalLock(lhDIB)
    If lpDIB = 0 Then GoTo ErrPlace
    
    Call CopyMemory(VarPtr(tBIH), lpDIB, Len(tBIH))
    
    lDIBSize = Len(tBIH) + (tBIH.biClrUsed * Len(tRGB)) + _
               (((tBIH.biWidth * tBIH.biBitCount + 31) \ 32) * 4 * tBIH.biHeight)
    ReDim bDIBits(1 To lDIBSize) As Byte
    Call CopyMemory(VarPtr(bDIBits(1)), lpDIB, lDIBSize)
    
    lRtn = GlobalUnlock(lhDIB)
    lRtn = GlobalFree(lhDIB)
    lhDIB = 0
    
    With tBFH
        .bfType = 19778     ' "BM"
        .bfSize = Len(tBFH) + lDIBSize
        .bfOffBits = Len(tBFH) + Len(tBIH) + (tBIH.biClrUsed * Len(tRGB))
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

Private Function TwainSetOneValue(ByVal Cap As TWAIN_MDL_CAPABILITY, _
                                  ByVal ItemType As TWAIN_MDL_ITEMYPE, _
                                  ByRef Item As Variant) As Long

    '-----------------------------------------------------------------------
    ' There are four types of containers that TWAIN defines for capabilities
    ' (TW_ONEVALUE, TW_ARRAY, TW_RANGE and TW_ENUMERATION)
    ' This module deals with one of them only - TW_ONEVALUE (single value)
    ' To set some capability you have to fill TW_ONEVALUE fields and use
    '   the triplet DG_CONTROL DAT_CAPABILITY MSG_SET
    ' The macros that convert some data types are used here as well
    '-----------------------------------------------------------------------
    On Local Error GoTo ErrPlace
    
    Dim tCapability As TW_CAPABILITY
    Dim tOneValue As TW_ONEVALUE
    Dim lhOneValue As Long
    Dim lpOneValue As Long
    Dim lRtn As Long
    Dim iRtn As Integer
    Dim tFix32 As TW_FIX32
    Dim iTmp As Integer

    tCapability.ConType = TWON_ONEVALUE
    tCapability.Cap = Cap
    
    tOneValue.ItemType = ItemType
    
    Select Case ItemType
        Case INT16
            tOneValue.Item1 = CInt(Item)
        Case UINT16, BOOL
            If ToUnsignedShort(CLng(Item), iTmp) Then GoTo ErrPlace
            Call CopyMemory(VarPtr(tOneValue.Item1), VarPtr(iTmp), 2&)
        Case FIX32
            If FloatToFix32(CSng(Item), tFix32) Then GoTo ErrPlace
            Call CopyMemory(VarPtr(tOneValue.Item1), VarPtr(tFix32), 4&)
    End Select

    lhOneValue = GlobalAlloc(GHND, Len(tOneValue))
    lpOneValue = GlobalLock(lhOneValue)
    Call CopyMemory(lpOneValue, VarPtr(tOneValue), Len(tOneValue))
    lRtn = GlobalUnlock(lhOneValue)
    tCapability.hContainer = lhOneValue

    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_CAPABILITY, MSG_SET, _
                     tCapability)
    If iRtn <> TWRC_SUCCESS Then
        lRtn = GlobalFree(lhOneValue)
        GoTo ErrPlace
    End If
    lRtn = GlobalFree(lhOneValue)

    TwainSetOneValue = 0
    Exit Function

ErrPlace:
    TwainSetOneValue = 1
End Function

Private Function TwainGetOneValue(ByVal Cap As TWAIN_MDL_CAPABILITY, _
                                  ByRef Item As Variant) As Long

    '-----------------------------------------------------------------------
    ' There are four types of containers that TWAIN defines for capabilities
    ' (TW_ONEVALUE, TW_ARRAY, TW_RANGE and TW_ENUMERATION)
    ' This module deals with one of them only - TW_ONEVALUE (single value)
    ' To get some capability you have to fill TW_ONEVALUE fields and use
    '   the triplet DG_CONTROL DAT_CAPABILITY MSG_GET
    ' The macros that convert some data types are used here as well
    '-----------------------------------------------------------------------
    
    On Local Error GoTo ErrPlace
    
    Dim tCapability As TW_CAPABILITY
    Dim tOneValue As TW_ONEVALUE
    Dim tFix32 As TW_FIX32
    Dim lpOneValue As Long
    Dim lRtn As Long
    Dim iRtn As Integer

    tCapability.ConType = TWON_ONEVALUE
    tCapability.Cap = Cap
    
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_CAPABILITY, MSG_GET, _
                     tCapability)
    If iRtn <> TWRC_SUCCESS Then GoTo ErrPlace
    
    lpOneValue = GlobalLock(tCapability.hContainer)
    Call CopyMemory(VarPtr(tOneValue), lpOneValue, Len(tOneValue))
    lRtn = GlobalUnlock(tCapability.hContainer)
    lRtn = GlobalFree(tCapability.hContainer)

    Select Case tOneValue.ItemType
        Case INT16
            Item = tOneValue.Item1
        Case UINT16, BOOL
            Item = FromUnsignedShort(tOneValue.Item1)
        Case FIX32
            Call CopyMemory(VarPtr(tFix32), VarPtr(tOneValue.Item1), 4&)
            Item = Fix32ToFloat(tFix32)
    End Select
    
    TwainGetOneValue = 0
    Exit Function

ErrPlace:
    TwainGetOneValue = 1
End Function

Private Function ToUnsignedShort(ByRef lSrc As Long, _
                                 ByRef iDst As Integer) As Long
    
    '------------------------------------------------------------------------
    ' Sets number ranging from 0 to 65535 into 2-byte VB Integer
    ' (useful for communicating with other dll that uses unsigned data types)
    '
    ' Function returns 0 is OK, 1 if an error occurs
    '------------------------------------------------------------------------
    
    On Local Error GoTo ErrPlace
    
    If (lSrc < 0) Or (lSrc > 65535) Then GoTo ErrPlace
    
    Call CopyMemory(VarPtr(iDst), VarPtr(lSrc), 2&)
    
    ' Another way
    'iDst = IIf(lSrc > 32767, lSrc - 65536, lSrc)
    
    ToUnsignedShort = 0
    Exit Function
    
ErrPlace:
    ToUnsignedShort = 1
End Function

Private Function FromUnsignedShort(ByRef iSrc As Integer) As Long

    '------------------------------------------------------------------------
    ' Gets the 2-byte unsigned number from VB Integer data type
    ' (useful for communicating with other dll that uses unsigned data types)
    '
    ' Function returns unsigned 2-byte value (in VB Long type)
    '------------------------------------------------------------------------
    
    Dim lTmp As Long
    
    Call CopyMemory(VarPtr(lTmp), VarPtr(iSrc), 2&)
    
    ' Another way
    'lTmp = IIf(iSrc < 0, iSrc + 65536, iSrc)
    
    FromUnsignedShort = lTmp

End Function

Private Function ToUnsignedLong(ByRef sngSrc As Single, _
                                ByRef lDst As Long) As Long

    '------------------------------------------------------------------------
    ' Sets number ranging from 0 to 4294967295 into 4-byte VB Long
    ' (useful for communicating with other dll that uses unsigned data types)
    '
    ' Function returns 0 is OK, 1 if an error occurs
    '------------------------------------------------------------------------

    On Local Error GoTo ErrPlace
    
    If (sngSrc < 0) Or (sngSrc > 4294967295#) Then GoTo ErrPlace
    
    lDst = IIf(sngSrc > 2147483647, sngSrc - 4294967296#, sngSrc)
    
    ToUnsignedLong = 0
    Exit Function
    
ErrPlace:
    ToUnsignedLong = 1
End Function

Private Function FromUnsignedLong(ByRef lSrc As Long) As Single

    '------------------------------------------------------------------------
    ' Gets the 4-byte unsigned number from VB Long data type
    ' (useful for communicating with other dll that uses unsigned data types)
    '
    ' Function returns unsigned 4-byte value (in VB Single type)
    '------------------------------------------------------------------------

    Dim sngTmp As Single
    
    sngTmp = IIf(lSrc < 0, lSrc + 4294967296#, lSrc)
    FromUnsignedLong = sngTmp

End Function

Private Function Fix32ToFloat(ByRef tFix32 As TW_FIX32) As Single
        
    '----------------------------------------------------------------
    ' Converts TWAIN TW_FIX32 data structure into VB Single data type
    ' (needed for communicating with TWAIN)
    '
    ' Function returns floating-point number in VB Single data type
    '----------------------------------------------------------------
        
    Dim sngTmp As Single
    
    sngTmp = tFix32.Whole + CSng(FromUnsignedShort(tFix32.Frac) / 65536)
    Fix32ToFloat = sngTmp

End Function

Private Function FloatToFix32(ByRef sngSrc As Single, _
                              ByRef tFix32 As TW_FIX32) As Long
    
    '----------------------------------------------------------------
    ' Converts VB Single data type into TWAIN TW_FIX32 data structure
    ' (needed for communicating with TWAIN)
    '
    ' Function returns 0 is OK, 1 if an error occurs
    '----------------------------------------------------------------
    
    On Local Error GoTo ErrPlace
    
    tFix32.Whole = CInt(Fix(sngSrc))
    Call ToUnsignedShort(CLng(sngSrc * 65536) And 65535, tFix32.Frac)
    FloatToFix32 = 0
    Exit Function

ErrPlace:
    FloatToFix32 = 1
End Function
