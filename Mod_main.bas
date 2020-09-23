Attribute VB_Name = "Mod_main"
Option Explicit
'======================================================================================
'welcome to this VB 3D Software engine
'
'This code show how to write custom Software 3D Renderer
'without DirectX, OpenGl, and any other API
'it uses only api for speeding up the drawing code
'
'He begins from the level Zero to pixel ploting over screen
'
'What does this code do?
'- Matrix Calculus and Transformation
'- Perspective correcting
'- Double Buffering
'- FPS Camera mouvement
'- Simple polygons rasterizing
'- .....and so...................
'
'This code is a start code for a full and complete Software 3D Renderer
'it can be converted to a Full Software 3D Engine
'
'IF YOU WANT INFO OR YOU WANT TO EMIT FEEDBACK
'mail me at johna_pop@yahoo.fr
'==========================================================================================

Function InitEngine(ByVal WindowHandle As Long, Optional doubleBuffering As Boolean = True) As Boolean

    GetClientRect WindowHandle, DATA.Buffer_Rect

    DATA.Buffer_Width = DATA.Buffer_Rect.Right - DATA.Buffer_Rect.Left
    DATA.Buffer_Height = DATA.Buffer_Rect.Bottom - DATA.Buffer_Rect.Top
    DATA.Buffer_Handle = WindowHandle
    DATA.Is_DoubleBuffering = doubleBuffering

    'we create the MatView matrix to its identity
    Matrix_Identity DATA.MatView

    'we create a brush to clear backbuffer
    SetEngine_ClearColor RGB(255, 255, 255)

    'here we prepare the backbuffer DC and bitmap

    ' We store the window DC
    DATA.Buffer_HDC = GetDC(DATA.Buffer_Handle)   ' ' Get the window's HDC

    ' We create a backbuffer compatible with then window current DC
    'so that we will be able to draw over
    DATA.Buffer_BackHDC = CreateCompatibleDC(DATA.Buffer_HDC)   '

    ' No worries, we'll just create a compatible bitmap the SAME size as
    ' our window's client area
    DATA.Buffer_Bitmap = CreateCompatibleBitmap(DATA.Buffer_HDC, DATA.Buffer_Width, DATA.Buffer_Height) '

    ' select the created back buffer into our back buffer's device context
    DATA.Buffer_OldBitmap = SelectObject(DATA.Buffer_BackHDC, DATA.Buffer_Bitmap)  '

    InitEngine = True

End Function

Sub Flip()

    GetClientRect DATA.Buffer_Handle, DATA.Buffer_Rect

    DATA.Buffer_Width = DATA.Buffer_Rect.Right - DATA.Buffer_Rect.Left
    DATA.Buffer_Height = DATA.Buffer_Rect.Bottom - DATA.Buffer_Rect.Top

    If DATA.Is_DoubleBuffering Then
        Call BitBlt(DATA.Buffer_HDC, 0, 0, DATA.Buffer_Rect.Right, _
             DATA.Buffer_Rect.Bottom, DATA.Buffer_BackHDC, 0, 0, vbSrcCopy)
    End If

End Sub

'===================================================================
' Here we clear the Buffer befor drawing any pixel
'
'====================================================================
Sub ClearBuffer()

  Dim DC As Long

  'we compute FPS here

    ComputeFPS

    'we make sure we clear the proper buffer
    DC = DATA.Buffer_HDC
    If DATA.Is_DoubleBuffering Then
        DC = DATA.Buffer_BackHDC
    End If

    'we fill the buffer with the color selected by user
    'clearcolor canbe changed by callin
    'SetEngine_ClearColor RGB(Red, Green, Blue)

    FillRect DC, DATA.Buffer_Rect, DATA.Buffer_ClearColorBrush

End Sub

'===================================================================
' Here we release any ressource create and used
' by the engine
'====================================================================
Sub Engine_Free()

    If DATA.Buffer_Handle = 0 Then Exit Sub

    If (DATA.Buffer_OldBitmap <> 0) Then ' If we have a double buffer

        ' Select back the original BITMAP
        Call SelectObject(DATA.Buffer_BackHDC, DATA.Buffer_OldBitmap)

        ' Free up memory
        Call DeleteObject(DATA.Buffer_Bitmap)
        Call DeleteDC(DATA.Buffer_BackHDC)
    End If

    If (DATA.Buffer_HDC <> 0) Then ' If the window's device context is valid

        ' Relese the device context
        Call ReleaseDC(DATA.Buffer_Handle, DATA.Buffer_HDC)
    End If

End Sub

'===================================================================
' Here we set the clear color
'====================================================================
Sub SetEngine_ClearColor(ByVal RGBcolor As Long)

    DATA.Buffer_ClearColor = RGBcolor
    DATA.Buffer_ClearColorBrush = CreateSolidBrush(DATA.Buffer_ClearColor)

End Sub

'===================================================================
' Draw a text over the buffer at x and y
'====================================================================
Sub Engine_DrawText(ByVal x As Long, ByVal y As Long, ByVal srcString As String)

    If DATA.Is_DoubleBuffering Then
        TextOut DATA.Buffer_BackHDC, x, y, srcString, Len(srcString)

      Else
        TextOut DATA.Buffer_HDC, x, y, srcString, Len(srcString)
    End If

End Sub

'=====================================================================================
' here we move the camera
'
'===================================================================================
Sub UpdateCamera()

   Dim Kam As VBSE_Matrix
  'if player has pressed UP key we move +Z
  If GetKeyPressed(vbKeyUp) Then DATA.Camera_Position.z = DATA.Camera_Position.z + 100 * DATA.TimeElapsed
  'if player has pressed DOWN key we move -Z
  If GetKeyPressed(vbKeyDown) Then DATA.Camera_Position.z = DATA.Camera_Position.z - 100 * DATA.TimeElapsed
  
  'we rotate camera
  If GetKeyPressed(vbKeyLeft) Then DATA.Camera_Rotation.y = DATA.Camera_Rotation.y - 0.5 * DATA.TimeElapsed
  If GetKeyPressed(vbKeyRight) Then DATA.Camera_Rotation.y = DATA.Camera_Rotation.y + 0.5 * DATA.TimeElapsed
  

  'finaly we update camera matrix here
  Matrix_Rotate_Y DATA.MatView, DATA.Camera_Rotation.y
  Matrix_Translate Kam, DATA.Camera_Position.x, DATA.Camera_Position.y, DATA.Camera_Position.z
  Matrix_Multiply Kam, DATA.MatView, DATA.MatView
        


End Sub

Sub GameLoop()

 
    'first we init the engine
    'we use doublebuffering
    InitEngine Form1.hwnd, True
 
    'here is an object that will be use
    'for our cube
    Dim CUBE1 As VBSE_Object
    Dim CUBE2 As VBSE_Object
   
    Dim M As VBSE_Matrix
   
   
        'first we create boxes
        Object_AddBox CUBE1, Vec3(-5, -5, -5), Vec3(5, 5, 5)
        Object_AddBox CUBE2, Vec3(-5, -5, -5), Vec3(5, 20, 5)
       
    'we translate the camera behind the objects
    DATA.Camera_Position = Vec3(0, 0, -150)
    

    Do
        'we let windows handle messages
        DoEvents
        
        'we handle player input
        UpdateCamera
        
        'we clear the buffer
        ClearBuffer
        'we rotate our 1st cube object
        Matrix_Rotate_Y CUBE1.WorldMatrix, Timer
        Draw_Object CUBE1, DATA.MatView, VBSE_LINE
        
        'we draw the left oject rotatin along X axis
        Matrix_Translate CUBE2.WorldMatrix, -20, 0, 0
        Matrix_Rotate_X M, -Timer
        Matrix_Multiply M, CUBE2.WorldMatrix, CUBE2.WorldMatrix
        Draw_Object CUBE2, DATA.MatView, VBSE_POINT
        
        'Here we do matrix operation
        'tranlate
        'rotate
        'scale
        'finalmat=tranlate * rotate * scale
         'we draw the left oject rotating along Y axis
        Matrix_Translate CUBE2.WorldMatrix, 50, 0, 100
        Matrix_Rotate_Y M, -Timer
        Matrix_Multiply M, CUBE2.WorldMatrix, CUBE2.WorldMatrix
        Matrix_Scale M, 1, 2, 1
        Matrix_Multiply M, CUBE2.WorldMatrix, CUBE2.WorldMatrix
        Draw_Object CUBE2, DATA.MatView, VBSE_LINE

        'we check if player does not want to quit
        If GetKeyPressed(vbKeyEscape) Then Call GameEnd

        Engine_DrawText 10, 10, "FPS=" + CStr(DATA.framesPerSecond)
        Engine_DrawText 10, 30, "Use Up/down to move, left/Right to rotate camera"
        Engine_DrawText 10, 45, "Press ESC to quit"

        Flip
    Loop

End Sub

Sub GameEnd()

    Engine_Free
    End

End Sub

Function GetKeyPressed(ByVal KEY As KeyCodeConstants) As Boolean

    GetKeyPressed = GetAsyncKeyState(KEY)

End Function

'===========================================================================================
'
'This is the core of the Rendering Pipeline
'
'here we convert a single 3D point to a 2D point
'
'
'
'
'============================================================================================

Sub Draw_Object(OBJ As VBSE_Object, CameraMatrix As VBSE_Matrix, Optional ByVal DrawinStyle As eVBSE_DRAWING = VBSE_LINE)

  Dim x As Single, y As Single, z As Single, tx As Single, ty As Single, tz As Single
  Dim num As Integer, I As Single, j As Single
  Dim M As VBSE_Matrix
  Dim lpPoint As POINTAPI
  Dim FirstPOINT As POINTAPI
  Dim LastPOINT As POINTAPI
  Dim NumPoint As Integer
  Dim DC As Long
  'for triangle drawing
  Dim Tris(1 To 3) As POINTAPI
  Dim Tris_Num As Integer

  'we select the proper buffer DC

    DC = DATA.Buffer_HDC
    If DATA.Is_DoubleBuffering Then
        DC = DATA.Buffer_BackHDC
    End If

  Dim bdrawLine As Boolean

    'first we compute the camera Inverse ViewMatrix
    Call Matrix_Inverse(CameraMatrix, M)

    'we concatenate it to the oject Worldmatrix
    Call Matrix_Multiply(M, OBJ.WorldMatrix, M)

    For num = 0 To OBJ.Num_Vertex - 1

        x = OBJ.Vertex(num).x
        y = OBJ.Vertex(num).y
        z = OBJ.Vertex(num).z

        'now we transform it from 3D to 2D
        tx = x * M.MM(0, 0) + y * M.MM(1, 0) + z * M.MM(2, 0) + M.MM(3, 0)
        ty = x * M.MM(0, 1) + y * M.MM(1, 1) + z * M.MM(2, 1) + M.MM(3, 1)
        tz = x * M.MM(0, 2) + y * M.MM(1, 2) + z * M.MM(2, 2) + M.MM(3, 2)
        '// transformace

        If (tz > 0.1) Then
            'we project the 3D to 2D
            I = (DATA.Buffer_Width / 2 + (500 * tx) / tz)
            j = (DATA.Buffer_Height / 2 - (500 * ty) / tz)

            If DrawinStyle = VBSE_TRIANGLE Then
                'we fill each 3 vertex to a triangle
    
                Tris_Num = Tris_Num + 1
                Tris(Tris_Num).x = I
                Tris(Tris_Num).y = j
                If Tris_Num >= 3 Then Tris_Num = 0
            End If
            

            If (I >= 0 And I < DATA.Buffer_Width And j >= 0 And j < DATA.Buffer_Height) Then
                NumPoint = NumPoint + 1

                'we store First and last point
                If NumPoint = 1 Then
                    FirstPOINT.x = I
                    FirstPOINT.y = j

                End If

                If DrawinStyle = VBSE_POINT Then

                    SetPixelV DC, I, j, 0
                    'we add this to draw large size of dot
                    SetPixelV DC, I + 1, j, 0
                    SetPixelV DC, I - 1, j, 0
                    SetPixelV DC, I, j + 1, 0
                    SetPixelV DC, I, j - 1, 0
                    SetPixelV DC, I + 1, j + 1, 0
                    SetPixelV DC, I - 1, j - 1, 0
                    SetPixelV DC, I - 1, j + 1, 0
                    SetPixelV DC, I + 1, j - 1, 0

                  ElseIf DrawinStyle = VBSE_LINE Then
                    If NumPoint > 1 Then
                        MoveToEx DC, I, j, lpPoint

                        LineTo DC, LastPOINT.x, LastPOINT.y

                    End If

                  ElseIf DrawinStyle = VBSE_TRIANGLE Then
                    'we draw the 1st line
                    MoveToEx DC, Tris(1).x, Tris(1).y, lpPoint
                    LineTo DC, Tris(2).x, Tris(2).y

                    'we draw the 2nd line
                    MoveToEx DC, Tris(2).x, Tris(2).y, lpPoint
                    LineTo DC, Tris(3).x, Tris(3).y

                    'we draw the 3rd line
                    MoveToEx DC, Tris(3).x, Tris(3).y, lpPoint
                    LineTo DC, Tris(1).x, Tris(1).y

                    'Polyline DC, Tris(1), 3

                End If

                LastPOINT.x = I
                LastPOINT.y = j

            End If

        End If

    Next num

 
End Sub

'===========================================================================================
'
'This sub add a box to an oject
'note the box has 2 tris per face then 36 points
'==========================================================================================
Function Object_AddBox(OBJ As VBSE_Object, Vmin As VBSE_Vector, Vmax As VBSE_Vector)

  Dim I As Integer

    I = OBJ.Num_Vertex
    OBJ.Num_Vertex = OBJ.Num_Vertex + 36
    ReDim Preserve OBJ.Vertex(OBJ.Num_Vertex - 1)

    'front
    OBJ.Vertex(I + 0) = Vec3(Vmin.x, Vmin.y, Vmin.z)
    OBJ.Vertex(I + 1) = Vec3(Vmin.x, Vmax.y, Vmin.z)
    OBJ.Vertex(I + 2) = Vec3(Vmax.x, Vmax.y, Vmin.z)
    OBJ.Vertex(I + 3) = Vec3(Vmax.x, Vmax.y, Vmin.z)
    OBJ.Vertex(I + 4) = Vec3(Vmax.x, Vmin.y, Vmin.z)
    OBJ.Vertex(I + 5) = Vec3(Vmin.x, Vmin.y, Vmin.z)
    'back
    OBJ.Vertex(I + 6) = Vec3(Vmax.x, Vmax.y, Vmax.z)
    OBJ.Vertex(I + 7) = Vec3(Vmin.x, Vmax.y, Vmax.z)
    OBJ.Vertex(I + 8) = Vec3(Vmin.x, Vmin.y, Vmax.z)
    OBJ.Vertex(I + 9) = Vec3(Vmin.x, Vmin.y, Vmax.z)
    OBJ.Vertex(I + 10) = Vec3(Vmax.x, Vmin.y, Vmax.z)
    OBJ.Vertex(I + 11) = Vec3(Vmax.x, Vmax.y, Vmax.z)
    'left
    OBJ.Vertex(I + 12) = Vec3(Vmax.x, Vmax.y, Vmax.z)
    OBJ.Vertex(I + 13) = Vec3(Vmax.x, Vmin.y, Vmax.z)
    OBJ.Vertex(I + 14) = Vec3(Vmax.x, Vmin.y, Vmin.z)
    OBJ.Vertex(I + 15) = Vec3(Vmax.x, Vmin.y, Vmin.z)
    OBJ.Vertex(I + 16) = Vec3(Vmax.x, Vmax.y, Vmin.z)
    OBJ.Vertex(I + 17) = Vec3(Vmax.x, Vmax.y, Vmax.z)
    'right
    OBJ.Vertex(I + 18) = Vec3(Vmin.x, Vmin.y, Vmin.z)
    OBJ.Vertex(I + 19) = Vec3(Vmin.x, Vmin.y, Vmax.z)
    OBJ.Vertex(I + 20) = Vec3(Vmin.x, Vmax.y, Vmax.z)
    OBJ.Vertex(I + 21) = Vec3(Vmin.x, Vmax.y, Vmax.z)
    OBJ.Vertex(I + 22) = Vec3(Vmin.x, Vmax.y, Vmin.z)
    OBJ.Vertex(I + 23) = Vec3(Vmin.x, Vmin.y, Vmin.z)
    'top
    OBJ.Vertex(I + 24) = Vec3(Vmin.x, Vmax.y, Vmin.z)
    OBJ.Vertex(I + 25) = Vec3(Vmin.x, Vmax.y, Vmax.z)
    OBJ.Vertex(I + 26) = Vec3(Vmax.x, Vmax.y, Vmax.z)
    OBJ.Vertex(I + 27) = Vec3(Vmax.x, Vmax.y, Vmax.z)
    OBJ.Vertex(I + 28) = Vec3(Vmax.x, Vmax.y, Vmin.z)
    OBJ.Vertex(I + 29) = Vec3(Vmin.x, Vmax.y, Vmin.z)
    'bottom
    OBJ.Vertex(I + 30) = Vec3(Vmax.x, Vmin.y, Vmax.z)
    OBJ.Vertex(I + 31) = Vec3(Vmin.x, Vmin.y, Vmax.z)
    OBJ.Vertex(I + 32) = Vec3(Vmin.x, Vmin.y, Vmin.z)
    OBJ.Vertex(I + 33) = Vec3(Vmin.x, Vmin.y, Vmin.z)
    OBJ.Vertex(I + 34) = Vec3(Vmax.x, Vmin.y, Vmin.z)
    OBJ.Vertex(I + 35) = Vec3(Vmax.x, Vmin.y, Vmax.z)

End Function

Private Sub ComputeFPS()

  'we convert milliseconds in sec

    DATA.currentTime = timeGetTime() * 0.001

    '    '' Here we store the elapsed time between the current and last frame,
    '    '' then keep the current frame in our static variable for the next frame.
    '
    DATA.NBframes = DATA.NBframes + 1

    DATA.TimeElapsed = DATA.currentTime - DATA.frameTime
    DATA.frameTime = DATA.currentTime

    '' Increase the frame counter
    DATA.framesPerSecond_counter = DATA.framesPerSecond_counter + 1

    '' Now we want to subtract the current time by the last time that was stored
    '' to see if the time elapsed has been over a second, which means we found our FPS.
    If (DATA.currentTime - DATA.LastTime > 1) Then

        '' Here we set the lastTime to the currentTime
        DATA.LastTime = DATA.currentTime

        '' Reset the frames per second
        DATA.framesPerSecond = DATA.framesPerSecond_counter

        DATA.framesPerSecond_counter = 0

    End If

End Sub
