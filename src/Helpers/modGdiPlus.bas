Attribute VB_Name = "modGdiPlus"
'@Folder("TPD_Addin.Helpers")

'===========================================================
'  LoadPictureGDIP - load any raster image file (PNG / JPG /
'  GIF / BMP / TIFF) into an stdole.IPictureDisp for an
'  MSForms Image control. VBA's built-in LoadPicture can't
'  read PNG; GDI+ can, and it needs no display context so it
'  works headlessly.
'
'  This is the only Declare in the codebase. It is isolated
'  here, guarded for VBA7 / Win64, and every failure path
'  returns Nothing (never raises) so callers can fall back to
'  a placeholder. Used by the Set TPD Defaults logo preview
'  ([#109](https://github.com/srjohnson1986/TPD-Addin-XLAM/issues/109)).
'
'  Standard Bullen/Pearson-style GDI+ picture loader.
'===========================================================

Option Explicit

#If VBA7 Then

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type PictDesc
    Size As Long
    Type As Long
    hPic As LongPtr
    hPal As LongPtr
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" ( _
    ByRef token As LongPtr, ByRef inputbuf As GdiplusStartupInput, ByVal outputbuf As LongPtr) As Long
Private Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" ( _
    ByVal fileName As LongPtr, ByRef bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" ( _
    ByVal bitmap As LongPtr, ByRef hbmReturn As LongPtr, ByVal background As Long) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" ( _
    ByRef PicDesc As PictDesc, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Long, _
    ByRef IPic As Object) As Long

#End If


' Returns the image at fileName as an IPictureDisp, or Nothing if the file is
' missing / unreadable / not a supported raster format, or on pre-VBA7 Office.
Public Function LoadPictureGDIP(ByVal fileName As String) As stdole.IPictureDisp

#If VBA7 Then
    Dim gsi As GdiplusStartupInput
    Dim token As LongPtr
    Dim gdipBitmap As LongPtr
    Dim hBitmap As LongPtr
    Dim pd As PictDesc
    Dim iidIPicture As GUID
    Dim newPic As Object

    If Len(Dir$(fileName)) = 0 Then Exit Function

    gsi.GdiplusVersion = 1
    If GdiplusStartup(token, gsi, 0) <> 0 Then Exit Function

    If GdipCreateBitmapFromFile(StrPtr(fileName), gdipBitmap) = 0 Then
        GdipCreateHBITMAPFromBitmap gdipBitmap, hBitmap, 0
        GdipDisposeImage gdipBitmap
    End If
    GdiplusShutdown token

    If hBitmap = 0 Then Exit Function

    With pd
        .Size = LenB(pd)
        .Type = 1               ' PICTYPE_BITMAP
        .hPic = hBitmap
        .hPal = 0
    End With

    ' IID_IPicture {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    With iidIPicture
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B: .Data4(1) = &HBB: .Data4(2) = &H0: .Data4(3) = &HAA
        .Data4(4) = &H0: .Data4(5) = &H30: .Data4(6) = &HC: .Data4(7) = &HAB
    End With

    If OleCreatePictureIndirect(pd, iidIPicture, 1, newPic) = 0 Then
        Set LoadPictureGDIP = newPic
    End If
#End If

End Function
