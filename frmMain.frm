VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Side Scrolling Example"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then Quit = True
If KeyCode = vbKeyRight Then sX = sX + 5
If KeyCode = vbKeyLeft Then sX = sX - 5
If KeyCode = vbKeyUp Then sY = sY - 5
If KeyCode = vbKeyDown Then sY = sY + 5

End Sub

Private Sub Form_Load()

bHandle = LoadSprite("sky.bmp")
oSky.SetBitmap (bHandle)

bHandle = LoadSprite("ground.bmp")
oGround.SetBitmap (bHandle)

bHandle = LoadSprite("tree.bmp")
oTree.SetBitmap (bHandle)

bHandle = LoadSprite("dactyl.bmp")
oCharacter.SetBitmap (bHandle)

Me.Width = (oSky.Width * Screen.TwipsPerPixelX)
Me.Height = oSky.Height * Screen.TwipsPerPixelY
Me.Refresh

Sky = 1
Ground = 6
Tree = 11
Frame = 0
Quit = False
    
End Sub
