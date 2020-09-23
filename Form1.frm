VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   265.735
   ScaleMode       =   0  'User
   ScaleWidth      =   377.072
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'
'
'Note: this is the first episode for a long serie
'so there is not any texture mapping, lighting, zbuffer, code.
'Enjoy this......................................
'==========================================================================================


Private Sub Form_Load()
Me.Caption = ENGINE_NAME + " " + ENGINE_VERION


Me.Refresh
Me.Show
Call GameLoop

End Sub

Private Sub Form_Unload(Cancel As Integer)
 GameEnd
End Sub
