Attribute VB_Name = "WindowCapture"
Option Explicit
' Test harness only: capture the native dialog including its title bar and OK button.
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type PictureDescription
    Size As Long
    Kind As Long
    Bitmap As Long
    Palette As Long
    Reserved As Long
End Type
Private Type Guid
    Part1 As Long
    Part2 As Long
    Part3 As Long
    Part4 As Long
End Type
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal index As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, bounds As Rect) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal dc As Long) As Long
Private Declare Function PrintWindow Lib "user32" (ByVal hwnd As Long, ByVal dc As Long, ByVal flags As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal dc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal dc As Long, ByVal width As Long, ByVal height As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal dc As Long, ByVal obj As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal dc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal obj As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (description As PictureDescription, iid As Guid, ByVal ownsHandle As Long, picture As IPicture) As Long

Public Sub SaveWindowImage(hwnd As Long, path As String)
    Dim bounds As Rect, source As Long, target As Long, bitmap As Long, previous As Long
    Dim desc As PictureDescription, iid As Guid, picture As IPicture, rendered As Long, result As Long
    Call GetWindowRect(hwnd, bounds)
    source = GetWindowDC(hwnd)
    target = CreateCompatibleDC(source)
    bitmap = CreateCompatibleBitmap(source, bounds.Right - bounds.Left, bounds.Bottom - bounds.Top)
    previous = SelectObject(target, bitmap)
    rendered = PrintWindow(hwnd, target, 0)
    Call SelectObject(target, previous)
    Call DeleteDC(target)
    Call ReleaseDC(hwnd, source)
    If rendered = 0 Then
        Call DeleteObject(bitmap)
        Err.Raise 5, , "Native window capture failed"
    End If
    desc.Size = Len(desc): desc.Kind = 1: desc.Bitmap = bitmap
    iid.Part1 = &H7BF80980: iid.Part2 = &H101ABF32: iid.Part3 = &HAA00BB8B: iid.Part4 = &HAB0C3000
    result = OleCreatePictureIndirect(desc, iid, 1, picture)
    If result <> 0 Then
        Call DeleteObject(bitmap)
        Err.Raise 5, , "Native bitmap conversion failed"
    End If
    SavePicture picture, path
    Set picture = Nothing
End Sub
