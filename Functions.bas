Attribute VB_Name = "Functions"
Function LoadSprite(sFilename As String) As Long

Dim hBitmap As Long     ' handle to bitmap being loaded
Dim oName As New clsbitmap

' use the LoadImage API to load in a bitmap image. The name of the
' image file is handed in the function
sFilename = App.Path & "\" & sFilename

' call the LoadImage API to attempt to load in the bitmap.
hBitmap = LoadImage(0, sFilename, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)

' if LoadImage fails, the return result is a zero, test for this
' before attempting to create our bitmap object.
If (hBitmap = 0) Then
    MsgBox "Error : Unable To Load Bitmap Image : " & sFilename, _
            vbOKOnly, "Bitmap Load Error"
    Exit Function
End If

LoadSprite = hBitmap

End Function
Public Function Draw()
BitBlt frmMain.hdc, 0, 0, oSky.Width, frmMain.Height, oSky.ImageDC, 0 + i * Sky, 0, SRCAND
BitBlt frmMain.hdc, 0, 0, oSky.Width, frmMain.Height, oSky.InvertImageDC, 0 + i * Sky, 0, SRCPAINT

    BitBlt frmMain.hdc, oSky.Width - i * Sky, 0, oSky.Width, frmMain.Height, oSky.ImageDC, 0, 0, SRCAND
    BitBlt frmMain.hdc, oSky.Width - i * Sky, 0, oSky.Width, frmMain.Height, oSky.InvertImageDC, 0, 0, SRCPAINT

BitBlt frmMain.hdc, 0, 0, oGround.Width, oGround.Height, oGround.ImageDC, 0 + (ii * Ground), 0, SRCAND
BitBlt frmMain.hdc, 0, 0, oGround.Width, oGround.Height, oGround.InvertImageDC, 0 + (ii * Ground), 0, SRCPAINT

    BitBlt frmMain.hdc, oGround.Width - ii * Ground, 0, oGround.Width, oGround.Height, oGround.ImageDC, 0, 0, SRCAND
    BitBlt frmMain.hdc, oGround.Width - ii * Ground, 0, oGround.Width, oGround.Height, oGround.InvertImageDC, 0, 0, SRCPAINT

BitBlt frmMain.hdc, 150 - iii * Tree, 100, oTree.Width, oTree.Height, oTree.ImageDC, 0, 0, SRCAND
BitBlt frmMain.hdc, 150 - iii * Tree, 100, oTree.Width, oTree.Height, oTree.InvertImageDC, 0, 0, SRCPAINT

BitBlt frmMain.hdc, 180 + sX, 100 + sY, 107, 60, oCharacter.ImageDC, 0 + (Frame * 107), 0, SRCAND
BitBlt frmMain.hdc, 180 + sX, 100 + sY, 107, 60, oCharacter.InvertImageDC, 0 + (Frame * 107), 0, SRCPAINT
frmMain.Refresh

End Function

