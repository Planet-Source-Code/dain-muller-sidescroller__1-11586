Attribute VB_Name = "Physics"
Public Function physics()
If droplul = True Then
        If jp <> 0 Then ys = ys - jp
        xs = xs + wind
        jp = jp - grav
        If jp < -jump Then droplul = False 'jp = jump '=Fix(Rnd * 30): jp = jump
End If
End Function

