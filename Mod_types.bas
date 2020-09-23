Attribute VB_Name = "mod_types"
'===================================================
'This module is used to the definition of all the data structures
'
'
'
'====================================================

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type POINTAPI
    x As Long
    y As Long
End Type

Public Const ENGINE_VERION = "0.01"
Public Const ENGINE_NAME = "Visual Basic SoftWare Engine"

'==============================================
'       Here is a matrix type
'       it will be used very often for
'       objects, MatView, transformations
'===========================================
Type VBSE_Matrix
    MM(3, 3) As Single
End Type

'==============================================
'       Here is a 3D vector type
'       it will be used very often for
'       Points, 3D vector algebra and so
'===========================================
Public Type VBSE_Vector
    x As Single
    y As Single
    z As Single
End Type

'==============================================
'       Here is a 3D Object
'
'===========================================
Public Type VBSE_Object
    'this will be used for Object transformation
    'in world space, like rotation,scaling, translation
    WorldMatrix As VBSE_Matrix

    'for object geometry datas
    'in fact any 3D oject
    'is formed by points
    Num_Vertex As Long
    'a list of points
    Vertex() As VBSE_Vector
End Type

Public Type VBSE_Scene

    Num_Objects As Long
    Objects() As VBSE_Object

End Type

Enum eVBSE_DRAWING
    VBSE_POINT
    VBSE_LINE
    VBSE_TRIANGLE
End Enum

'this is datas for engine
'information
Public Type Engine_data
    Buffer_Rect As RECT
    Buffer_HDC As Long
    Buffer_Width As Single
    Buffer_Height As Single

    Buffer_Handle As Long 'this is a copy of window handle

    ' ' This is the window's HDC
    Buffer_BackHDC As Long  ' ' this will be used for double Buffering

    Buffer_Bitmap As Long ' This is the HBITMAP that we'll fill with an exact copy of
    ' our window's HBITMAP (pixels)

    ' We need this guy around so we can totally free up memory when it's
    ' all said and done
    Buffer_OldBitmap  As Long

    Buffer_ClearColor As Long
    Buffer_ClearColorBrush As Long
    FPS As Long
    MatView As VBSE_Matrix

    Camera_Position As VBSE_Vector
    Camera_Rotation As VBSE_Vector

    World As VBSE_Scene

    'data for frame per second counter
    TimeElapsed As Single

    framesPerSecond_counter    As Single     '/ This will store our fps

    framesPerSecond    As Single     '/ This will store our fps
    LastTime          As Single     '' This will hold the time from the last frame
    currentTime As Single

    frameTime  As Single
    NBframes As Currency

    'flags
    Is_DoubleBuffering As Boolean

End Type

Public DATA As Engine_data

