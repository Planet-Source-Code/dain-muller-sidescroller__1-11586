Attribute VB_Name = "Loop"
Sub main()
frmMain.Show

Do
i = i + 1
ii = ii + 1
iii = iii + 1
Frame = Frame + 1
Call Draw
If Frame = 4 Then Frame = 0

If i * Sky >= oSky.Width Then i = 0
If ii * Ground >= oGround.Width Then ii = 0
If 120 - iii * Tree < -oTree.Width Then iii = -50


DoEvents
Loop Until Quit = True

If Quit = True Then
    Unload frmMain
    End
End If
End Sub
