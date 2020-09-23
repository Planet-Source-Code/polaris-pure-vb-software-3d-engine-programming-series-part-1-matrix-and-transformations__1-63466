Attribute VB_Name = "Mod_Math"

Option Explicit
'========================================================================================
' Some API S for faster drawing via GDI
'========================================================================================

'we use this to perform memory copy

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long

Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Declare Function timeGetTime Lib "winmm.dll" () As Long

Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long



'==================================================
'  Matrix Identity
'
'
'
'================================================

Sub Matrix_Identity(ByRef RetMat As VBSE_Matrix)

    
        RetMat.MM(0, 1 - 1) = 1
        RetMat.MM(0, 2 - 1) = 0
        RetMat.MM(0, 3 - 1) = 0
        RetMat.MM(0, 4 - 1) = 0
        RetMat.MM(1, 1 - 1) = 0
        RetMat.MM(1, 2 - 1) = 1
        RetMat.MM(1, 3 - 1) = 0
        RetMat.MM(1, 4 - 1) = 0
        RetMat.MM(2, 1 - 1) = 0
        RetMat.MM(2, 2 - 1) = 0
        RetMat.MM(2, 3 - 1) = 1
        RetMat.MM(2, 4 - 1) = 0
        RetMat.MM(3, 1 - 1) = 0
        RetMat.MM(3, 2 - 1) = 0
        RetMat.MM(3, 3 - 1) = 0
        RetMat.MM(3, 4 - 1) = 1

End Sub

Sub Matrix_Scale(M As VBSE_Matrix, ByVal x As Single, ByVal y As Single, ByVal z As Single)
 
 Matrix_Identity M
 
 M.MM(0, 0) = x
 M.MM(1, 1) = y
 M.MM(2, 2) = z
 
 
 
End Sub


'==================================================
'  Matrix Multiply
'
'  r = a * b
'
'================================================

Sub Matrix_Multiply(a As VBSE_Matrix, b As VBSE_Matrix, r As VBSE_Matrix)

  Dim temp As VBSE_Matrix
  Dim I As Integer, j As Integer

    For I = 0 To 3
        For j = 0 To 3
            temp.MM(I, j) = a.MM(0, j) * b.MM(I, 0) + a.MM(1, j) * b.MM(I, 1) + _
                    a.MM(2, j) * b.MM(I, 2) + a.MM(3, j) * b.MM(I, 3)
        Next j
    Next I

    r = temp

End Sub

'==================================================
'  Matrix Rotation along X axis
'
'
'
'================================================

Sub Matrix_Rotate_X(M As VBSE_Matrix, ByVal angle As Single)

  Dim temp As VBSE_Matrix

    Matrix_Identity temp
    temp.MM(2, 1) = -Sin(angle)
    temp.MM(2, 2) = Cos(angle)
    temp.MM(1, 1) = Cos(angle)
    temp.MM(1, 2) = Sin(angle)

    M = temp
End Sub


'==================================================
'  Matrix Rotation along Y axis
'
'
'
'================================================

Sub Matrix_Rotate_Y(M As VBSE_Matrix, ByVal angle As Single)

  Dim temp As VBSE_Matrix

    Matrix_Identity temp
    
    temp.MM(0, 0) = Cos(angle)
    temp.MM(0, 2) = -Sin(angle)
    temp.MM(2, 2) = Cos(angle)
    temp.MM(2, 0) = Sin(angle)

   M = temp

End Sub

'==================================================
'  Matrix Rotation along Y axis
'
'
'
'================================================
Sub Matrix_Rotate_Z(M As VBSE_Matrix, ByVal angle As Single)

  Dim temp As VBSE_Matrix

    Matrix_Identity temp
    
    temp.MM(0, 0) = Cos(angle)
    temp.MM(0, 1) = Sin(angle)
    temp.MM(1, 1) = Cos(angle)
    temp.MM(1, 0) = -Sin(angle)

  M = temp

End Sub




'==================================================
'  Matrix Translation along XYZ axes
'
'================================================
Sub Matrix_Translate(M As VBSE_Matrix, ByVal x As Single, ByVal y As Single, ByVal z As Single)

  Dim temp As VBSE_Matrix

    Matrix_Identity temp
    
    temp.MM(3, 0) = x
    temp.MM(3, 1) = y
    temp.MM(3, 2) = z
  

    M = temp

End Sub


'==================================================
'  MatRetert Matrix MatSrc and store the result in MatRet Matrix
'
'================================================
Sub Matrix_Inverse(MatSrc As VBSE_Matrix, MatRet As VBSE_Matrix)

Dim I As Integer, j As Integer
Dim f_det As Single

   f_det = MatSrc.MM(0, 0) * MatSrc.MM(1, 1) * MatSrc.MM(2, 2) + _
            MatSrc.MM(1, 0) * MatSrc.MM(2, 1) * MatSrc.MM(0, 2) + _
            MatSrc.MM(0, 1) * MatSrc.MM(1, 2) * MatSrc.MM(2, 0) - _
            MatSrc.MM(0, 2) * MatSrc.MM(1, 1) * MatSrc.MM(2, 0) - _
            MatSrc.MM(0, 1) * MatSrc.MM(1, 0) * MatSrc.MM(2, 2) - _
            MatSrc.MM(0, 0) * MatSrc.MM(1, 2) * MatSrc.MM(2, 1)
    
    If f_det = 0 Then Exit Sub

    MatRet.MM(0, 0) = MatSrc.MM(1, 1) * MatSrc.MM(2, 2) _
                   - MatSrc.MM(1, 2) * MatSrc.MM(2, 1)
                   
    MatRet.MM(1, 0) = -MatSrc.MM(1, 0) * MatSrc.MM(2, 2) + MatSrc.MM(1, 2) * MatSrc.MM(2, 0)
    
    
    MatRet.MM(2, 0) = MatSrc.MM(1, 0) * MatSrc.MM(2, 1) - MatSrc.MM(1, 1) * MatSrc.MM(2, 0)
        
    MatRet.MM(3, 0) = -MatSrc.MM(1, 0) * MatSrc.MM(2, 1) * MatSrc.MM(3, 2) - _
                   MatSrc.MM(1, 1) * MatSrc.MM(2, 2) * MatSrc.MM(3, 0) - _
                   MatSrc.MM(2, 0) * MatSrc.MM(3, 1) * MatSrc.MM(1, 2) + _
                   MatSrc.MM(1, 2) * MatSrc.MM(2, 1) * MatSrc.MM(3, 0) + _
                   MatSrc.MM(1, 1) * MatSrc.MM(2, 0) * MatSrc.MM(3, 2) + _
                   MatSrc.MM(2, 2) * MatSrc.MM(1, 0) * MatSrc.MM(3, 1)
  
  

    MatRet.MM(0, 1) = -MatSrc.MM(0, 1) * MatSrc.MM(2, 2) + MatSrc.MM(0, 2) * MatSrc.MM(2, 1)
    MatRet.MM(1, 1) = MatSrc.MM(0, 0) * MatSrc.MM(2, 2) - MatSrc.MM(0, 2) * MatSrc.MM(2, 0)
    MatRet.MM(2, 1) = -MatSrc.MM(0, 0) * MatSrc.MM(2, 1) + MatSrc.MM(0, 1) * MatSrc.MM(2, 0)
    
    MatRet.MM(3, 1) = MatSrc.MM(0, 0) * MatSrc.MM(2, 1) * MatSrc.MM(3, 2) + _
                   MatSrc.MM(0, 1) * MatSrc.MM(2, 2) * MatSrc.MM(3, 0) + _
                   MatSrc.MM(2, 0) * MatSrc.MM(3, 1) * MatSrc.MM(0, 2) - _
                   MatSrc.MM(0, 2) * MatSrc.MM(2, 1) * MatSrc.MM(3, 0) - _
                   MatSrc.MM(0, 1) * MatSrc.MM(2, 0) * MatSrc.MM(3, 2) - _
                   MatSrc.MM(2, 2) * MatSrc.MM(0, 0) * MatSrc.MM(3, 1)
                   
    MatRet.MM(0, 2) = MatSrc.MM(0, 1) * MatSrc.MM(1, 2) - MatSrc.MM(0, 2) * MatSrc.MM(1, 1)
    MatRet.MM(1, 2) = -MatSrc.MM(0, 0) * MatSrc.MM(1, 2) + MatSrc.MM(0, 2) * MatSrc.MM(1, 0)
    MatRet.MM(2, 2) = MatSrc.MM(0, 0) * MatSrc.MM(1, 1) - MatSrc.MM(0, 1) * MatSrc.MM(1, 0)
    
    MatRet.MM(3, 2) = -MatSrc.MM(0, 0) * MatSrc.MM(1, 1) * MatSrc.MM(3, 2) - _
                   MatSrc.MM(0, 1) * MatSrc.MM(1, 2) * MatSrc.MM(3, 0) - _
                   MatSrc.MM(1, 0) * MatSrc.MM(3, 1) * MatSrc.MM(0, 2) + _
                   MatSrc.MM(0, 2) * MatSrc.MM(1, 1) * MatSrc.MM(3, 0) + _
                   MatSrc.MM(0, 1) * MatSrc.MM(1, 0) * MatSrc.MM(3, 2) + _
                   MatSrc.MM(1, 2) * MatSrc.MM(0, 0) * MatSrc.MM(3, 1)

   For I = 0 To 3
        For j = 0 To 3
            MatRet.MM(j, I) = MatRet.MM(j, I) / f_det
        Next j
   Next I
   
    MatRet.MM(0, 3) = 0
    MatRet.MM(1, 3) = 0
    MatRet.MM(2, 3) = 0
    MatRet.MM(3, 3) = 1
    

End Sub

Sub Matrix_DebugPrint(M As VBSE_Matrix)
Dim I As Integer
Dim j As Integer
 For I = 0 To 3
        For j = 0 To 3
            Debug.Print "MM" + CStr(I + 1) & CStr(j + 1) & "=" + CStr(M.MM(I, j))
        Next j
   Next I

End Sub



'===========================================================================================
'  Vector 3D operations
'
'===========================================================================================

Function Vec3(ByVal x As Single, ByVal y As Single, ByVal z As Single) As VBSE_Vector

    Vec3.x = x
    Vec3.y = y
    Vec3.z = z
    
End Function




