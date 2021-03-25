# Accessing TWAIN compatible scanner in Visual Basic

## PREFACE
This page describes the way VB (5/6) programmers can access TWAIN scanners.

While developing some application I needed to access a scanner without showing the `UI (User Interface)` that comes with each scanner driver. I was searching the Internet to get some VB code concerning this topic but I found some DLLs or OCXs only (I must mention here a very handy dll called eztwain from [www.dosadi.com](http://www.dosadi.com)). Finally I decided to dig into that problem by myself.

I made a VB Module that can transfer one image at a time from scanner and save it as the `BMP (Bitmap Image file)` file. Later on, this VB Module was enhanced with the `ADF (Automatic Document Feeder)` feature. There are three public functions in the module:

- **PopupSelectSourceDialog** - shows TWAIN dialog for selecting default data source for acquisition.
- **TransferWithoutUI** - transfers image(s) from TWAIN data source without showing the data source UI (silent transfer). The programmer can set following parameters of the transfer:
  - Image resolution (DPI).
  - Image colour depth - monochromatic, grey, fullcolour.
  - Image size and position - left, top, right, bottom (in inches).
  - Acquisition source - scanner glass or ADF.

  The image(s) is saved as the BMP file.
- **TransferWithUI** - transfers image(s) from TWAIN data source using the data source UI to set image and transfer parameters.\
  The image(s) is saved as the BMP file.

Functions return 0 if everything is OK, 1 if an error occurs.

## ABOUT TWAIN
The way one can access an optical scanner is through a standard called TWAIN. If you want to get real knowledge about it I strongly recommend you to read the twain specification located in [www.twain.org](https://www.twain.org). If you just want to touch it briefly, I can tell you that your application can communicate with the scanner using the `DSM_Entry` function located in the twain library.

To simply get an image one must:

- open data source manager (twain dll)
- open data source (scanner driver)
- enable data source (with visible or hidden UI)
- transfer image from data source
- disable data source
- close data source
- close data source manager.

Using different TWAIN triplets as the DSM_Entry function parameters, the programmer communicates with the data source manager or with the data source. For example a triplet `DG_CONTROL DAT_PARENT MSG_OPENDSM` opens data source manager, a triplet `DG_CONTROL DAT_PARENT MSG_CLOSEDSM` closes data source manager.

## DECLARATION
First of all one must convert the C declaration from twain.h into the VB declaration. This is not much difficult if one has experienced it before. Declaration in twain.h says that TW_UINT32 is an `unsigned long` type, TW_UINT16 is an `unsigned short` type, pTW_IDENTITY is a `pointer to TW_IDENTITY` structure and TW_MEMREF is a `pointer to void`.

Therefore
```
TW_UINT16 FAR PASCAL DSM_Entry (pTW_IDENTITY pOrigin,
                                pTW_IDENTITY pDest,
                                TW_UINT32    DG,
                                TW_UINT16    DAT,
                                TW_UINT16    MSG,
                                TW_MEMREF    pData);
```
is converted into
```
Private Declare Function DSM_Entry Lib "TWAIN_32.DLL" ( _
                                   ByRef pOrigin As Any, _
                                   ByRef pDest As Any, _
                                   ByVal DG As Long, _
                                   ByVal DAT As Integer, _
                                   ByVal MSG As Integer, _
                                   ByRef pData As Any) As Integer
```
But there is a problem with converting some `UDT (User Defined Type)` and it is called the `byte alignment`. VB uses 4-byte alignment in UDTs and adds so-called `padding bytes` to keep this kind of alignment, while C uses 2-byte alignment. If `Len` of UDT differs from `LenB` of UDT, you know there are some padding bytes in this UDT. Here in TWAIN, it is important to eliminate those padding bytes in order to get VB values into right places (bytes) in the C type UDT.

Luckily, you can use the `Win32Api CopyMemory` function. The example below shows how to put the VB Long value into TW_ONEVALUE.Item, using the CopyMemory function and the pointer to UDT:

_Example_

C declaration
```
typedef struct {
    TW_UINT16  ItemType;
    TW_UINT32  Item;
} TW_ONEVALUE, FAR * pTW_ONEVALUE;
```
VB declaration
```
Private Type TW_ONEVALUE
    ItemType As Integer               ' TW_UINT16
    Item     As Long                  ' TW_UINT32
End Type
```
VB code to assign the VB Long value to TW_ONEVALUE.Item
```
Dim tOneValue As TW_ONEVALUE
Dim lValue As Long

lValue = 1354
Call CopyMemory(VarPtr(tOneValue) + 2, VarPtr(lValue), 4&)
```
If you assigned value to TW_ONEVALUE.Item using the VB standard way `tOneValue.Item = lValue`, it would be shifted two bytes right due to 2 padding bytes placed after TW_ONEVALUE.ItemType. When C program receives such UDT, it cannot get the right value from the UDT variable Item. The CopyMemory function makes sure the value of lValue is put correctly into Item variable of the C type UDT.

## EXAMPLES OF USE
Finally, there are some examples of how to use this module in your code. Please note that now you can use ADF, which means you can acquire more than one image during a scanning session. Therefore you need to provide the file counter by reference. The module will use this counter when naming the acquired image(s) and increment it after each successful transfer. You also need to provide the image file name, which the module automatically adds the suffix .bmp to.

1. Let us scan a monochromatic 100 DPI image from the scanner glass and save it as `noui_bw_<counter>.bmp`. We want to scan the whole scanner glass (set sngImageRight and sngImageBottom to 0) and we want to use silent transfer without showing the data source UI. ADF is disabled:
```
Dim lRtn As Long
lRtn = mdlTwain.TransferWithoutUI(100, BW, 0, 0, 0, 0, _
                                  "noui_bw", lCounter, False)
```
2. Let us scan greyscale 200 DPI images from ADF and save them as `noui_grey_<counter>.bmp`. We want to scan the whole ADF acquire area (set sngImageRight and sngImageBottom to zero) and we want to use silent transfer without showing the data source UI. ADF is enabled:
```
Dim lRtn As Long
lRtn = mdlTwain.TransferWithoutUI(200, GREY, 0, 0, 0, 0, _
                                  "noui_grey", lCounter, True)
```
3. Let us scan a colour 300 DPI image from the scanner glass and save it as `noui_colour_<counter>.bmp`. We want to scan the partial rectangle (left=1, top=1, right=2, bottom=5 inches) of the scanner glass and we want to use silent transfer without showing the data source UI. ADF is disabled:
```
Dim lRtn As Long
lRtn = mdlTwain.TransferWithoutUI(300, RGB, 1, 1, 2, 5, _
                                  "noui_colour", lCounter, False)
````
4. Let us scan image(s) using the data source UI to set the attributes and save it as `ui_img_<counter>.bmp`:
```
Dim lRtn As Long
lRtn = mdlTwain.TransferWithUI("ui_img", lCounter)
```
5. Let us choose the default data source (scanner) that will be used for next transfers:
```
Dim lRtn As Long
lRtn = mdlTwain.PopupSelectSourceDialog()
```

## UPDATES
- 2021/03/21 - Project moved to GitHub
- 2020/03/15 - Version 1.1.
  - ADF (Automatic Document Feeder) capability implemented.
  - A few bug fixes.
- 2004/10/14 - Major changes. I have completely rebuilt the code and have numbered it as version 1.0.
  - As I left my first ideas about this code to be a class, I have changed it into VB module.
  - Added option to transfer image with UI enabled - new public function TransferWithUI.
  - Added option to select source - new public function PopupSelectSourceDialog.
  - Added universal functions to set and get transfer and image attributes.
  - Added some macros for converting data type.
  - Added a lot of comments (hope useful).
- 2004/03/30 - The new class version with corrected DMS_Entry function declaration.

## CONTACT
I appreciate your feedback, you can send it to [lumir.mik@gmail.com](mailto:lumir.mik@gmail.com).

Let me say a word of thanks to Alfred Koppold, who helped me with setting the scanner resolution and whose feedback helped me to understand the byte alignment problem.

[![Donate with Paypal](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=5SEP8ZE5PXZSN)
