Attribute VB_Name = "MMat"
Option Explicit
Public Type Point3
    X As Double
    Y As Double
    Z As Double
End Type
Public Type Point4
    X As Double
    Y As Double
    Z As Double
    W As Double
End Type
Public Type Matrix23 '2 Zeilen 3 Spalten
    aa As Double 'Zeile 1 Spalte 1
    ab As Double 'Zeile 1 Spalte 2
    ac As Double 'Zeile 1 Spalte 3
    ba As Double 'Zeile 2 Spalte 1
    bb As Double 'Zeile 2 Spalte 2
    bc As Double 'Zeile 2 Spalte 3
End Type

'[aa  ab]
'[ba  bb]
'[ca  cb]

'[X] = [aa ab ac] * [X]   [aa*X+ab*Y+ac*Z]
'[Y]   [ba bb bc]   [Y] = [ba*X+bb*Y+bc*Z]
'                   [Z]
Public Type Matrix33
    aa As Double 'Zeile 1 Spalte 1
    ab As Double 'Zeile 1 Spalte 2
    ac As Double 'Zeile 1 Spalte 3
    ba As Double 'Zeile 2 Spalte 1
    bb As Double 'Zeile 2 Spalte 2
    bc As Double 'Zeile 2 Spalte 3
    ca As Double 'Zeile 3 Spalte 1
    CB As Double 'Zeile 3 Spalte 2
    cc As Double 'Zeile 3 Spalte 3
End Type
'[aa  ab  ac]   [11 12 13]   [aa*11+ab*21+ac*31 aa*12+ab*22+ac*32 aa*13+ab*23+ac*33]
'[ba  bb  bc] * [21 22 23] = [ba*11+bb*21+bc*31 ba*12+bb*22+bc*32 ba*13+bb*23+bc*33]
'[ca  cb  cc]   [31 32 33]   [ca*11+cb*21+cc*31 ca*12+cb*22+cc*32 ca*13+cb*23+cc*33]

Public Type Matrix34 '3 Zeilen 4 Spalten
    aa As Double 'Zeile 1 Spalte 1
    ab As Double 'Zeile 1 Spalte 2
    ac As Double 'Zeile 1 Spalte 3
    ad As Double 'Zeile 1 Spalte 4
    ba As Double 'Zeile 2 Spalte 1
    bb As Double 'Zeile 2 Spalte 2
    bc As Double 'Zeile 2 Spalte 3
    bd As Double 'Zeile 2 Spalte 4
    ca As Double 'Zeile 3 Spalte 1
    CB As Double 'Zeile 3 Spalte 2
    cc As Double 'Zeile 3 Spalte 3
    cd As Double 'Zeile 3 Spalte 4
End Type

Public Type Matrix44
    aa As Double 'Zeile 1 Spalte 1
    ab As Double 'Zeile 1 Spalte 2
    ac As Double 'Zeile 1 Spalte 3
    ad As Double 'Zeile 1 Spalte 4
    ba As Double 'Zeile 2 Spalte 1
    bb As Double 'Zeile 2 Spalte 2
    bc As Double 'Zeile 2 Spalte 3
    bd As Double 'Zeile 2 Spalte 4
    ca As Double 'Zeile 3 Spalte 1
    CB As Double 'Zeile 3 Spalte 2
    cc As Double 'Zeile 3 Spalte 3
    cd As Double 'Zeile 3 Spalte 4
    da As Double 'Zeile 4 Spalte 1
    db As Double 'Zeile 4 Spalte 2
    dc As Double 'Zeile 4 Spalte 3
    dd As Double 'Zeile 4 Spalte 4
End Type

'Public Function New_Point2(ByVal X As Long, ByVal Y As Long) As Point2L
'    With New_TPoint2
'        .X = X
'        .Y = Y
'    End With
'End Function
Public Function New_Point3(ByVal X As Double, ByVal Y As Double, ByVal Z As Double) As Point3
    With New_Point3
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function
Public Function New_Point4(ByVal X As Double, ByVal Y As Double, ByVal Z As Double, ByVal W As Double) As Point4
    With New_Point4
        .X = X
        .Y = Y
        .Z = Z
        .W = W
    End With
End Function

Public Function New_Matrix23(ByVal aa As Double, ByVal ab As Double, ByVal ac As Double, _
                             ByVal ba As Double, ByVal bb As Double, ByVal bc As Double) As Matrix23
    With New_Matrix23
        .aa = aa:        .ab = ab:        .ac = ac:
        .ba = ba:        .bb = bb:        .bc = bc:
    End With
End Function

Public Function New_Matrix33(ByVal aa As Double, ByVal ab As Double, ByVal ac As Double, _
                             ByVal ba As Double, ByVal bb As Double, ByVal bc As Double, _
                             ByVal ca As Double, ByVal CB As Double, ByVal cc As Double) As Matrix33
    With New_Matrix33
        .aa = aa:         .ab = ab:         .ac = ac
        .ba = ba:         .bb = bb:         .bc = bc
        .ca = ca:         .CB = CB:         .cc = cc
    End With
End Function
Public Function New_Matrix34(ByVal aa As Double, ByVal ab As Double, ByVal ac As Double, ByVal ad As Double, _
                             ByVal ba As Double, ByVal bb As Double, ByVal bc As Double, ByVal bd As Double, _
                             ByVal ca As Double, ByVal CB As Double, ByVal cc As Double, ByVal cd As Double) As Matrix34
    With New_Matrix34
        .aa = aa:         .ab = ab:         .ac = ac:         .ad = ad
        .ba = ba:         .bb = bb:         .bc = bc:         .bd = bd
        .ca = ca:         .CB = CB:         .cc = cc:         .cd = cd
    End With
End Function

Public Function New_Matrix44(ByVal aa As Double, ByVal ab As Double, ByVal ac As Double, ByVal ad As Double, _
                             ByVal ba As Double, ByVal bb As Double, ByVal bc As Double, ByVal bd As Double, _
                             ByVal ca As Double, ByVal CB As Double, ByVal cc As Double, ByVal cd As Double, _
                             ByVal da As Double, ByVal db As Double, ByVal dc As Double, ByVal dd As Double) As Matrix44
    With New_Matrix44
        .aa = aa:         .ab = ab:         .ac = ac:         .ad = ad
        .ba = ba:         .bb = bb:         .bc = bc:         .bd = bd
        .ca = ca:         .CB = CB:         .cc = cc:         .cd = cd
        .da = da:         .db = db:         .dc = dc:         .dd = dd
    End With
End Function
Public Function New_Matrix44get33(mat As Matrix33) As Matrix44
    With New_Matrix44get33
        .aa = mat.aa: .ab = mat.ab: .ac = mat.ac
        .ba = mat.ba: .bb = mat.bb: .bc = mat.bc
        .ca = mat.ca: .CB = mat.CB: .cc = mat.cc
    End With
End Function

Public Function Mat33Multiply(mat1 As Matrix33, mat2 As Matrix33) As Matrix33
    With Mat33Multiply
        .aa = mat1.aa * mat2.aa + mat1.ab * mat2.ba + mat1.ac * mat2.ca:        .ab = mat1.aa * mat2.ab + mat1.ab * mat2.bb + mat1.ac * mat2.CB:        .ac = mat1.aa * mat2.ac + mat1.ab * mat2.bc + mat1.ac * mat2.cc
        .ba = mat1.ba * mat2.aa + mat1.bb * mat2.ba + mat1.bc * mat2.ca:        .bb = mat1.ba * mat2.ab + mat1.bb * mat2.bb + mat1.bc * mat2.CB:        .bc = mat1.ba * mat2.ac + mat1.bb * mat2.bc + mat1.bc * mat2.cc
        .ca = mat1.ca * mat2.aa + mat1.CB * mat2.ba + mat1.cc * mat2.ca:        .CB = mat1.ca * mat2.ab + mat1.CB * mat2.bb + mat1.cc * mat2.CB:        .cc = mat1.ca * mat2.ac + mat1.CB * mat2.bc + mat1.cc * mat2.cc
    End With
End Function

Public Function Mat44Multiply(mat1 As Matrix44, mat2 As Matrix44) As Matrix44
    With Mat44Multiply
        .aa = mat1.aa * mat2.aa + mat1.ab * mat2.ba + mat1.ac * mat2.ca + mat1.ad * mat2.da:        .ab = mat1.aa * mat2.ab + mat1.ab * mat2.bb + mat1.ac * mat2.CB + mat1.ad * mat2.db:        .ac = mat1.aa * mat2.ac + mat1.ab * mat2.bc + mat1.ac * mat2.cc + mat1.ad * mat2.dc:        .ad = mat1.aa * mat2.ad + mat1.ab * mat2.bd + mat1.ac * mat2.cd + mat1.ad * mat2.dd
        .ba = mat1.ba * mat2.aa + mat1.bb * mat2.ba + mat1.bc * mat2.ca + mat1.bd * mat2.da:        .bb = mat1.ba * mat2.ab + mat1.bb * mat2.bb + mat1.bc * mat2.CB + mat1.bd * mat2.db:        .bc = mat1.ba * mat2.ac + mat1.bb * mat2.bc + mat1.bc * mat2.cc + mat1.bd * mat2.dc:        .bd = mat1.ba * mat2.ad + mat1.bb * mat2.bd + mat1.bc * mat2.cd + mat1.bd * mat2.dd
        .ca = mat1.ca * mat2.aa + mat1.CB * mat2.ba + mat1.cc * mat2.ca + mat1.cd * mat2.da:        .CB = mat1.ca * mat2.ab + mat1.CB * mat2.bb + mat1.cc * mat2.CB + mat1.cd * mat2.db:        .cc = mat1.ca * mat2.ac + mat1.CB * mat2.bc + mat1.cc * mat2.cc + mat1.cd * mat2.dc:        .cd = mat1.ca * mat2.ad + mat1.CB * mat2.bd + mat1.cc * mat2.cd + mat1.cd * mat2.dd
        .da = mat1.da * mat2.aa + mat1.db * mat2.ba + mat1.dc * mat2.ca + mat1.dd * mat2.da:        .db = mat1.da * mat2.ab + mat1.db * mat2.bb + mat1.dc * mat2.CB + mat1.dd * mat2.db:        .dc = mat1.da * mat2.ac + mat1.db * mat2.bc + mat1.dc * mat2.cc + mat1.dd * mat2.dc:        .dd = mat1.da * mat2.ad + mat1.db * mat2.bd + mat1.dc * mat2.cd + mat1.dd * mat2.dd
    End With
End Function
Public Function Mat34Multiply44(mat1 As Matrix34, mat2 As Matrix44) As Matrix34
    With Mat34Multiply44
        .aa = mat1.aa * mat2.aa + mat1.ab * mat2.ba + mat1.ac * mat2.ca + mat1.ad * mat2.da:        .ab = mat1.aa * mat2.ab + mat1.ab * mat2.bb + mat1.ac * mat2.CB + mat1.ad * mat2.db:        .ac = mat1.aa * mat2.ac + mat1.ab * mat2.bc + mat1.ac * mat2.cc + mat1.ad * mat2.dc:        .ad = mat1.aa * mat2.ad + mat1.ab * mat2.bd + mat1.ac * mat2.cd + mat1.ad * mat2.dd
        .ba = mat1.ba * mat2.aa + mat1.bb * mat2.ba + mat1.bc * mat2.ca + mat1.bd * mat2.da:        .bb = mat1.ba * mat2.ab + mat1.bb * mat2.bb + mat1.bc * mat2.CB + mat1.bd * mat2.db:        .bc = mat1.ba * mat2.ac + mat1.bb * mat2.bc + mat1.bc * mat2.cc + mat1.bd * mat2.dc:        .bd = mat1.ba * mat2.ad + mat1.bb * mat2.bd + mat1.bc * mat2.cd + mat1.bd * mat2.dd
        .ca = mat1.ca * mat2.aa + mat1.CB * mat2.ba + mat1.cc * mat2.ca + mat1.cd * mat2.da:        .CB = mat1.ca * mat2.ab + mat1.CB * mat2.bb + mat1.cc * mat2.CB + mat1.cd * mat2.db:        .cc = mat1.ca * mat2.ac + mat1.CB * mat2.bc + mat1.cc * mat2.cc + mat1.cd * mat2.dc:        .cd = mat1.ca * mat2.ad + mat1.CB * mat2.bd + mat1.cc * mat2.cd + mat1.cd * mat2.dd
        '.da = mat1.da * mat2.aa + mat1.db * mat2.ba + mat1.dc * mat2.ca + mat1.dd * mat2.da:        .db = mat1.da * mat2.ab + mat1.db * mat2.bb + mat1.dc * mat2.cb + mat1.dd * mat2.db:        .dc = mat1.da * mat2.ac + mat1.db * mat2.bc + mat1.dc * mat2.cc + mat1.dd * mat2.dc:        .dd = mat1.da * mat2.ad + mat1.db * mat2.bd + mat1.dc * mat2.cd + mat1.dd * mat2.dd
    End With
End Function

Public Function MatMul_Mat34_Pt4(aMat As Matrix34, aPt As Point4) As Point3
    With MatMul_Mat34_Pt4
        .X = aMat.aa * aPt.X + aMat.ab * aPt.Y + aMat.ac * aPt.Z + aMat.ad * aPt.W
        .Y = aMat.ba * aPt.X + aMat.bb * aPt.Y + aMat.bc * aPt.Z + aMat.bd * aPt.W
        .Z = aMat.ca * aPt.X + aMat.CB * aPt.Y + aMat.cc * aPt.Z + aMat.cd * aPt.W
    End With
End Function

Public Function Matrix33_ToStr(aMat As Matrix33) As String
    With aMat
        Matrix33_ToStr = "[" & .aa & " " & .ab & " " & .ac & "]" & vbCrLf & _
                         "[" & .ba & " " & .bb & " " & .bc & "]" & vbCrLf & _
                         "[" & .ca & " " & .CB & " " & .cc & "]" & vbCrLf

    End With
End Function
Public Function Matrix34_ToStr(aMat As Matrix34) As String
    With aMat
        Matrix34_ToStr = "[" & .aa & " " & .ab & " " & .ac & " " & .ad & "]" & vbCrLf & _
                         "[" & .ba & " " & .bb & " " & .bc & " " & .bd & "]" & vbCrLf & _
                         "[" & .ca & " " & .CB & " " & .cc & " " & .cd & "]" & vbCrLf
    End With
End Function
Public Function Matrix44_ToStr(aMat As Matrix44) As String
    With aMat
        Matrix44_ToStr = "[" & .aa & " " & .ab & " " & .ac & " " & .ad & "]" & vbCrLf & _
                         "[" & .ba & " " & .bb & " " & .bc & " " & .bd & "]" & vbCrLf & _
                         "[" & .ca & " " & .CB & " " & .cc & " " & .cd & "]" & vbCrLf & _
                         "[" & .da & " " & .db & " " & .dc & " " & .dd & "]" & vbCrLf
    End With
End Function

Public Function Matrix33_AbsNorm(aMat As Matrix33) As Double
    Dim v As Double
    With aMat
        v = .aa * .aa + .ab * .ab + .ac * .ac _
          + .ba * .ba + .bb * .bb + .bc * .bc _
          + .ca * .ca + .CB * .CB + .cc * .cc
    End With
    Matrix33_AbsNorm = VBA.Math.Sqr(v)
End Function

Public Function New_Matrix33RotX(ByVal alp_X As Double) As Matrix33
    Dim SinAlp_X As Double: SinAlp_X = VBA.Math.Sin(alp_X)
    Dim CosAlp_X As Double: CosAlp_X = IIf(alp_X = Pi2, 0, VBA.Math.Cos(alp_X))
    New_Matrix33RotX = New_Matrix33(1, 0, 0, _
                                    0, CosAlp_X, -SinAlp_X, _
                                    0, SinAlp_X, CosAlp_X)
End Function
Public Function New_Matrix33RotY(ByVal bet_Y As Double) As Matrix33
    Dim SinBet_Y As Double: SinBet_Y = VBA.Math.Sin(bet_Y)
    Dim CosBet_Y As Double: CosBet_Y = IIf(bet_Y = Pi2, 0, VBA.Math.Cos(bet_Y))
    New_Matrix33RotY = New_Matrix33(CosBet_Y, 0, SinBet_Y, _
                                      0, 1, 0, _
                                      -SinBet_Y, 0, CosBet_Y)
End Function
Public Function New_Matrix33RotZ(ByVal gam_Z As Double) As Matrix33
    Dim SinGam_Z As Double: SinGam_Z = VBA.Math.Sin(gam_Z)
    Dim CosGam_Z As Double: CosGam_Z = IIf(gam_Z = Pi2, 0, VBA.Math.Cos(gam_Z))
    New_Matrix33RotZ = New_Matrix33(CosGam_Z, -SinGam_Z, 0, _
                                      SinGam_Z, CosGam_Z, 0, _
                                      0, 0, 1)
End Function
Public Function New_Matrix33RotXYZ(ByVal alp_X As Double, ByVal bet_Y As Double, ByVal gam_Z As Double) As Matrix33
    New_Matrix33RotXYZ = Mat33Multiply(New_Matrix33RotX(alp_X), Mat33Multiply(New_Matrix33RotY(bet_Y), New_Matrix33RotZ(gam_Z)))
End Function

'die Winkel im Bogenmass angeben
Public Function New_Matrix44RotXYZ(ByVal alp_X As Double, ByVal bet_Y As Double, ByVal gam_Z As Double) As Matrix44
'    Dim SinAlp_X As Double: SinAlp_X = VBA.Math.Sin(alp_X)
'    Dim CosAlp_X As Double: CosAlp_X = VBA.Math.Cos(alp_X)
'    Dim MatX As TMatrix33: MatX = New_TMatrix33(1, 0, 0, _
'                                                0, CosAlp_X, -SinAlp_X, _
'                                                0, SinAlp_X, CosAlp_X)
'    Dim SinBet_Y As Double: SinBet_Y = VBA.Math.Sin(bet_Y)
'    Dim CosBet_Y As Double: CosBet_Y = VBA.Math.Cos(bet_Y)
'    Dim MatY As TMatrix33: MatY = New_TMatrix33(CosBet_Y, 0, SinBet_Y, _
'                                                0, 1, 0, _
'                                                -SinBet_Y, 0, CosBet_Y)
'    Dim SinGam_Z As Double: SinGam_Z = VBA.Math.Sin(gam_Z)
'    Dim CosGam_Z As Double: CosGam_Z = VBA.Math.Cos(gam_Z)
'    Dim MatZ As TMatrix33: MatZ = New_TMatrix33(CosGam_Z, -SinGam_Z, 0, _
'                                                SinGam_Z, CosGam_Z, 0, _
'                                                0, 0, 1)
'    New_TMatrix33RotXYZ = MatMultiply33(MatMultiply33(MatX, MatY), MatZ)
    Dim tmp As Matrix33: tmp = New_Matrix33RotXYZ(alp_X, bet_Y, gam_Z)
    With New_Matrix44RotXYZ
        .aa = tmp.aa:        .ab = tmp.ab: .ac = tmp.ac
        .ba = tmp.ba:        .bb = tmp.bb: .bc = tmp.bc
        .ca = tmp.ca:        .CB = tmp.CB: .cc = tmp.cc
                                                         .dd = 1
    End With
End Function
'// Diese Funktion erstellt eine 4x4-Rotations-Matrix
'Mat_<Type> rotation_matrix(Type alpha, Type beta, Type gamma) {
'    // Rotation um X-Achse
'    Type r_x_data[] = {
'                        1,           0,           0,
'                        0,           +cos(alpha), -sin(alpha),
'                        0,           +sin(alpha), +cos(alpha)
'                      };
'    // Rotation um Y-Achse
'    Type r_y_data[] = {
'                        +cos(beta),  0,           +sin(beta),
'                        0,           1,           0,
'                        -sin(beta),  0,           +cos(beta)
'                      };
'    // Rotation um Z-Achse
'    Type r_z_data[] = {
'                        +cos(gamma), -sin(gamma), 0,
'                        +sin(gamma), +cos(gamma), 0,
'                        0,           0,           1
'                      };
'
'    // Rotationsmatrix erstellen
'    Mat_<Type> rotation(4, 4, 0.0);
'    rotation(3, 3) = 1;
'    Mat tmp = Mat_<Type>(3, 3, r_x_data) * Mat_<Type>(3, 3, r_y_data) * Mat_<Type>(3, 3, r_z_data);
'
'    Mat test = rotation.colRange(0, 3).rowRange(0, 3);
'    tmp.copyTo(test);
'
'    //tmp.copyTo((const cv::Mat)(rotation.colRange(0, 3).rowRange(0, 3)));
'
'    return rotation;
'};

'von hier:
'www.is.uni-due.de/fileadmin/lehre/13w_crv/uebung/06_Kamera.cpp
'z.B. w_b = 3000, s_x=s_y=1, b0x=0, b0y=0, ocx=0, ocy=0, ocz=15, ax=0, by=0, gz=0
Public Function New_Matrix34Camera(ByVal w_b As Double, _
                                   ByVal s_x As Double, ByVal s_y As Double, _
                                   ByVal B_x0 As Double, ByVal B_y0 As Double, _
                                   ByVal O_cx As Double, ByVal O_cy As Double, ByVal O_cz As Double, _
                                   ByVal alp_X As Double, ByVal bet_Y As Double, ByVal gam_Z As Double) As Matrix34
    Dim extr As Matrix44: extr = New_Matrix44RotXYZ(alp_X, bet_Y, gam_Z)
    With extr:        .ad = O_cx:        .bd = O_cy:        .cd = O_cz:    End With
    'Debug.Print Matrix44_ToStr(extr)
    Dim intr As Matrix34: intr = New_Matrix34(-w_b / s_x, 0, B_x0, 0, _
                                              0, -w_b / s_y, B_y0, 0, _
                                              0, 0, 1, 0)
    New_Matrix34Camera = Mat34Multiply44(intr, extr)
End Function
'
'// Diese Funktion erstellt eine Kamera-Matrix
'
'Mat_<Type> camera_matrix (
'    Type w_b,
'    Type s_x, Type s_y,
'    Type B_x0, Type B_y0,
'    Type O_cx, Type O_cy, Type O_cz,
'    Type alpha, Type beta, Type gamma
') {
'    // erstellen der extrinischen Matrix
'    Mat_<Type> extrinsic = rotation_matrix(alpha, beta, gamma);
'    Type translation_data[] = {
'                                O_cx,
'                                O_cy,
'                                O_cz
'                              };
'    Mat_<Type> translation(3, 1, translation_data);
'
'    Mat test = extrinsic.col(3).rowRange(0, 3);
'    translation.copyTo(test);
'    //translation.copyTo(extrinsic.col(3).rowRange(0, 3));
'
'    // erstellen der intrinsischen Matrix
'    Type intrinsic_data[] = {
'                                -w_b/s_x, 0,        B_x0, 0,
'                                0,        -w_b/s_y, B_y0, 0,
'                                0,        0,        1,    0
'                            };
'    Mat_<Type> intrinsic(3, 4, intrinsic_data);
'
'    return intrinsic*extrinsic;
'};

Public Function Point3_Projection(p As Point3, aMat As Matrix34) As Point2L
    Dim z2 As Double
    With aMat
        z2 = .ca * p.X + .CB * p.Y + .cc * p.Z + .cd * 1
        If z2 = 0 Then Exit Function
        Point3_Projection.X = (.aa * p.X + .ab * p.Y + .ac * p.Z + .ad * 1) / z2
        Point3_Projection.Y = (.ba * p.X + .bb * p.Y + .bc * p.Z + .bd * 1) / z2
    End With
End Function


