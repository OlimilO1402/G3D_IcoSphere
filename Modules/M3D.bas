Attribute VB_Name = "M3D"
Option Explicit
Public Type Triangle
    iP1 As Long 'index zum 1. Punkt des Dreiecks im Array aus Point3
    iP2 As Long 'index zum 2. Punkt des Dreiecks im Array aus Point3
    iP3 As Long 'index zum 3. Punkt des Dreiecks im Array aus Point3
    iNorm As Long 'index zur Flächennormalen
End Type
Public Type Object3D
    points()     As Point3   'die Punkte
    Triangles()  As Triangle 'Indices zu den Punkten
    normales()   As Point3   'die Flächennormalen der Triangles
    projection() As Point2L  'die Punkte in projizierten Bildkoordinaten
End Type
Public Const eps  As Double = 0.000000000001
Private HashTable As Collection
Public sc As Double

Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As Point2L, ByVal nCount As Long) As Long
        
Public Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As Point2L, ByVal nCount As Long) As Long

'Constructors
Public Function New_Point3(ByVal X As Double, ByVal Y As Double, ByVal Z As Double) As Point3

    With New_Point3:        .X = X:        .Y = Y:        .Z = Z:    End With
    
End Function

Public Function Point3_IsEqual(p1 As Point3, p2 As Point3) As Boolean
    Point3_IsEqual = (Abs(p1.X - p2.X) < eps) And (Abs(p1.Y - p2.Y) < eps) And (Abs(p1.Z - p2.Z) < eps)
End Function

Public Function Point3_Subtract(p1 As Point3, p2 As Point3) As Point3
    With Point3_Subtract
        .X = p1.X - p2.X:        .Y = p1.Y - p2.Y:        .Z = p1.Z - p2.Z
    End With
End Function
Public Function New_Point3_va(v) As Point3

    If Not IsArray(v) Then Exit Function
    Dim n As Long: n = UBound(v) - LBound(v) + 1
    
    With New_Point3_va
        .X = CDbl(v(0)): If n < 2 Then Exit Function
        .Y = CDbl(v(1)): If n < 3 Then Exit Function
        .Z = CDbl(v(2))
    End With
    
End Function

Public Function New_Triangle(ByVal i1 As Long, ByVal i2 As Long, ByVal i3 As Long) As Triangle

    With New_Triangle:        .iP1 = i1:        .iP2 = i2:        .iP3 = i3:    End With
    
End Function

Public Function New_Triangle_va(v) As Triangle

    If Not IsArray(v) Then Exit Function
    Dim n As Long: n = UBound(v) - LBound(v) + 1
    
    With New_Triangle_va
        .iP1 = CLng(v(0)): If n < 2 Then Exit Function
        .iP2 = CLng(v(1)): If n < 3 Then Exit Function
        .iP3 = CLng(v(2))
    End With
    
End Function

Public Function IcosahedronPoint_r(ByVal sphere_radius As Double) As Point3
    'sphere_radius = Umkugelradius
    IcosahedronPoint_r = IcosahedronPoint_a(sphere_radius * 4 / (Sqr(2 * (5 + Sqr(5)))))
End Function
Public Function IcosahedronPoint_a(ByVal edge_length As Double) As Point3
    'edgelength = Kantenlänge a
    Dim X As Double: X = edge_length / 2       '0.525731112119134
    Dim Z As Double: Z = X / 2 * (1 + Sqr(5))  '0.85065080835204
    IcosahedronPoint_a = New_Point3(X, 0, Z)
End Function

Public Function CheckEulerPolyeder(obj As Object3D) As String
    Dim s1 As String, s2 As String
    Dim e As Long, f As Long, k As Long
    With obj
        e = UBound(.points) - LBound(.points) + 1
        f = UBound(.Triangles) - LBound(.Triangles) + 1
    End With
    k = f * 3 / 2
    s1 = "Ecken: e; Flächen: f; Kanten: k" & vbCrLf & "e = " & e & "; f = " & f & "; k = " & k & vbCrLf & "Körper ist "
    s2 = "e + f = k + 2" & vbCrLf & e & " + " & f
    If e + f = k + 2 Then
        s2 = s2 & " = "
    Else
        s1 = s1 & "k"
        s2 = s2 & " <> "
    End If
    s1 = s1 & "ein Euler-Polyeder, da"
    s2 = s2 & k & " + 2"
    
    CheckEulerPolyeder = s1 & vbCrLf & s2
End Function
Public Function CreateIcosahedron(p As Point3) As Object3D
    'erzeugt ein Ikosaeder (engl: icosahedron)
    'die X und Z-Koordinaten des angegebenen Punktes werden verwendet
    Dim X As Double: X = Abs(p.X)
    Dim Z As Double: Z = Abs(p.Z)
    If X = 0 And Z = 0 Then X = 0.525731112119134: Z = 0.85065080835204
        
    Dim s1: s1 = v(v(-X, 0#, Z), v(X, 0#, Z), v(-X, 0#, -Z), v(X, 0#, -Z), _
                   v(0#, Z, X), v(0#, Z, -X), v(0#, -Z, X), v(0#, -Z, -X), _
                   v(Z, X, 0#), v(-Z, X, 0#), v(Z, -X, 0#), v(-Z, -X, 0#))
    Dim s2: s2 = v(v(0, 4, 1), v(0, 9, 4), v(9, 5, 4), v(4, 5, 8), v(4, 8, 1), _
                   v(8, 10, 1), v(8, 3, 10), v(5, 3, 8), v(5, 2, 3), v(2, 7, 3), _
                   v(7, 10, 3), v(7, 6, 10), v(7, 11, 6), v(11, 0, 6), v(0, 1, 6), _
                   v(6, 1, 10), v(9, 0, 11), v(9, 11, 2), v(9, 2, 5), v(7, 2, 11))
                   
    Dim icosphere As Object3D
    With icosphere
        .points = ParsePoints(s1)
        .Triangles = ParseTriangles(s2)
        'halt, so hat man zwar die Normalen aber noch nicht die Indizes in den Triangles
        '.normales = CreateNormales(icosphere)
    End With
    CreateNormales icosphere
    CreateIcosahedron = icosphere
End Function

Sub CreateNormales(obj As Object3D) 'As Point3()
    With obj
        Dim i As Long: i = LBound(.Triangles)
        Dim u As Long: u = UBound(.Triangles)
        ReDim .normales(i To u) As Point3
        Dim T As Triangle
        For i = i To u
            .Triangles(i).iNorm = i
            T = .Triangles(i)
            .normales(i) = TriangleNormale(.points(T.iP1), .points(T.iP2), .points(T.iP3))
        Next
    End With
End Sub
'Public Function Icosahedron_subdivide_old(obj As Object3D) As Object3D
'
'    Dim ics As Object3D: ics = obj
'    Dim i As Long
'    Dim t_i As Triangle
'    Dim v1  As Point3, v2  As Point3, v3  As Point3
'    Dim v12 As Point3, v23 As Point3, v31 As Point3
'    Dim i12   As Long, i23   As Long, i31 As Long
'    Dim t_l As Long: t_l = LBound(ics.Triangles)
'    Dim t_u As Long: t_u = UBound(ics.Triangles)
'    Dim p_l As Long: p_l = LBound(ics.Points)
'    Dim p_u As Long: p_u = UBound(ics.Points)
'
'    For i = t_l To t_u
'
'        t_i = ics.Triangles(i)
'        v1 = ics.Points(t_i.iP1)
'        v2 = ics.Points(t_i.iP2)
'        v3 = ics.Points(t_i.iP3)
'
'        'durch Subdivision der Dreiecke neue Punkte v12, v23, v31 gewinnen
'        '      v2
'        '      /\                 /\
'        '     /  \           v12 /__\ v23
'        '    /    \      =>     /\  /\
'        'v1 /______\ v3        /__\/__\
'        '                        v13
'        '
'        Subdivide v1, v2, v3, v12, v23, v31
'
'        'die neuen Punkte hinzufügen
'        i12 = p_u + 1:            i23 = p_u + 2:            i31 = p_u + 3
'        p_u = p_u + 3
'        ReDim Preserve ics.Points(p_l To p_u)
'        ics.Points(i12) = v12
'        ics.Points(i23) = v23
'        ics.Points(i31) = v31
'
'        'und die neuen Dreiecke hinzufügen
'        t_u = t_u + 4
'        ReDim Preserve ics.Triangles(t_l To t_u)
'        ics.Triangles(t_u - 3) = New_Triangle(t_i.iP1, i12, i31)
'        ics.Triangles(t_u - 2) = New_Triangle(t_i.iP2, i23, i12)
'        ics.Triangles(t_u - 1) = New_Triangle(t_i.iP3, i31, i23)
'        ics.Triangles(t_u - 0) = New_Triangle(i12, i23, i31)
'
'    Next
'    Icosahedron_subdivide = ics
'
'End Function
Private Sub HashTable_Add(p As Point3, i As Long) 'As Collection
    Dim H As String: H = Str(Round(p.X, 7)) & "|" & Str(Round(p.Y, 7)) & "|" & Str(Round(p.Z, 7))
    HashTable.Add i, H
    'Set CreateHash = HashTable
End Sub

Public Function HashTable_Contains(p As Point3) As Long
    'Ja nö, so is das wahrscheinlich ein Schmarrn, weil
    Dim H As String: H = Str(Round(p.X, 7)) & "|" & Str(Round(p.Y, 7)) & "|" & Str(Round(p.Z, 7))
    On Error Resume Next
    HashTable_Contains = CLng(HashTable(H))
    If Err.Number Then
        'enthält nicht
        HashTable_Contains = -1
    End If
    'If IsEmpty(HashTable(h)) Then: 'DoNothing
    'If (Err.Number = 0) Then
    '    ContainsPoint = CLng(HashTable(h))
    'End If
    On Error GoTo 0
End Function

Public Function Icosahedron_subdivide(obj As Object3D) As Object3D

'i: 0 | p:    12 | t:    20
'i: 1 | p:    42 | t:    80
'i: 2 | p:   162 | t:   320
'i: 3 | p:   642 | t:  1280
'i: 4 | p:  2562 | t:  5120
'i: 5 | p: 10242 | t: 20480
'i: 6 | p: 40962 | t: 81920

    Set HashTable = New Collection
    Dim ics As Object3D: ics = obj
    Dim i As Long, ii As Long
    Dim t_i As Triangle
    Dim v1  As Point3, v2  As Point3, v3  As Point3
    Dim v12 As Point3, v23 As Point3, v31 As Point3
    Dim i12   As Long, i23   As Long, i31 As Long
    Dim t_l As Long: t_l = LBound(ics.Triangles)
    Dim t_u As Long: t_u = UBound(ics.Triangles)
    Dim p_l As Long: p_l = LBound(ics.points)
    Dim p_u As Long: p_u = UBound(ics.points)
    
    For i = t_l To t_u
    
        t_i = ics.Triangles(i)
        v1 = ics.points(t_i.iP1)
        v2 = ics.points(t_i.iP2)
        v3 = ics.points(t_i.iP3)
        
        'durch Subdivision der Dreiecke neue Punkte v12, v23, v31 gewinnen
        '
        '      v2
        '      /\                 /\
        '     /  \           v12 /__\ v23
        '    /    \      =>     /\  /\
        'v1 /______\ v3        /__\/__\
        '                        v13
        '
        Subdivide v1, v2, v3, v12, v23, v31
        
        'die neuen Punkte hinzufügen
        ii = HashTable_Contains(v12)
        If ii > -1 Then i12 = ii Else p_u = p_u + 1: i12 = p_u: HashTable_Add v12, i12
        
        ii = HashTable_Contains(v23)
        If ii > -1 Then i23 = ii Else p_u = p_u + 1: i23 = p_u: HashTable_Add v23, i23
        
        ii = HashTable_Contains(v31)
        If ii > -1 Then i31 = ii Else p_u = p_u + 1: i31 = p_u: HashTable_Add v31, i31
        
        ReDim Preserve ics.points(p_l To p_u)
        ics.points(i12) = v12
        ics.points(i23) = v23
        ics.points(i31) = v31
        
        'und die neuen Dreiecke hinzufügen
        t_u = t_u + 3
        ReDim Preserve ics.Triangles(t_l To t_u)
        ics.Triangles(t_u - 2) = New_Triangle(t_i.iP1, i12, i31)
        ics.Triangles(t_u - 1) = New_Triangle(t_i.iP2, i23, i12)
        ics.Triangles(t_u - 0) = New_Triangle(t_i.iP3, i31, i23)
        ics.Triangles(i) = New_Triangle(i12, i23, i31)
        
    Next
    Icosahedron_subdivide = ics
        
End Function
'Public Function IsInCollection( _
'    ByRef col As Collection, _
'    ByRef elem As String _
'  ) As Boolean
'
'  On Error Resume Next
'
'    If IsEmpty(col(elem)) Then: 'DoNothing
'    IsInCollection = (Err.Number = 0)
'
'  On Error GoTo 0
'
'End Function
'Public Function ContainsPoint(points() As Point3, searchpt As Point3) As Long
'    'gibt den Index zurück, andernfalls -1
' dauert viel zu lang!!!
' wir brauchen eine HashTable
'    Dim i As Long
'    For i = LBound(points) To UBound(points)
'        If Point3_IsEqual(searchpt, points(i)) Then
'            ContainsPoint = i
'            Exit Function
'        End If
'    Next
'    ContainsPoint = -1
'End Function
Function Subdivide(v1 As Point3, v2 As Point3, v3 As Point3, out_v12 As Point3, out_v23 As Point3, out_v31 As Point3)
    
    With out_v12:        .X = v1.X + v2.X:        .Y = v1.Y + v2.Y:        .Z = v1.Z + v2.Z:    End With
    out_v12 = Normalize(out_v12)
    
    With out_v23:        .X = v2.X + v3.X:        .Y = v2.Y + v3.Y:        .Z = v2.Z + v3.Z:    End With
    out_v23 = Normalize(out_v23)
    
    With out_v31:        .X = v3.X + v1.X:        .Y = v3.Y + v1.Y:        .Z = v3.Z + v1.Z:    End With
    out_v31 = Normalize(out_v31)
End Function

Public Function Normalize(p As Point3) As Point3

    Dim d As Double: d = VBA.Sqr(p.X * p.X + p.Y * p.Y + p.Z * p.Z)
    If d = 0 Then Exit Function
    
    With Normalize
        .X = p.X / d
        .Y = p.Y / d
        .Z = p.Z / d
    End With
    
End Function
Public Function TriangleNormale(p1 As Point3, p2 As Point3, p3 As Point3) As Point3
    TriangleNormale = NormCrossProd(Point3_Subtract(p1, p2), Point3_Subtract(p2, p3))
End Function
Public Function NormCrossProd(v1 As Point3, v2 As Point3) As Point3
    With NormCrossProd
        .X = v1.Y * v2.Z - v1.Z * v2.Y
        .Y = v1.Z * v2.X - v1.X * v2.Z
        .Z = v1.X * v2.Y - v1.Y * v2.X
    End With
    NormCrossProd = Normalize(NormCrossProd)
End Function

'Hilfsfunktion
Public Function v(ParamArray p())
    v = p
End Function

'einen Polyeder eiförmig machen
Public Function Eggify_x(obj As Object3D, x_null As Double, x_fact_pos As Double, x_fact_neg As Double) As Object3D
    Dim egg As Object3D: egg = obj
    With egg
        Dim i As Long
        For i = LBound(.points) To UBound(.points)
            With .points(i)
                If .X > x_null Then
                    .X = .X * x_fact_pos
                Else
                    .X = .X * x_fact_neg
                End If
            End With
        Next
    End With
    Eggify_x = egg
End Function
Public Function Eggify_z(obj As Object3D, z_null As Double, z_fact_pos As Double, z_fact_neg As Double) As Object3D
    Dim egg As Object3D: egg = obj
    With egg
        Dim i As Long
        For i = LBound(.points) To UBound(.points)
            With .points(i)
                If .Z > z_null Then
                    .Z = .Z * z_fact_pos
                Else
                    .Z = .Z * z_fact_neg
                End If
            End With
        Next
    End With
    Eggify_z = egg
End Function

'Zeichenroutinen
Public Function DrawObj3D_xy(aPB As PictureBox, obj As Object3D, color As Long)

    Dim i As Long
    Dim T As Triangle
    Dim p1 As Point3, p2 As Point3, p3 As Point3
    'Dim sc As Double: sc = 37 * 5
    If sc = 0 Then sc = 37
    Dim tx As Double: tx = aPB.ScaleWidth / 2
    Dim ty As Double: ty = aPB.ScaleHeight / 2
    Dim c As Long
    aPB.ForeColor = color
    With obj
    
        For i = LBound(.Triangles) To UBound(.Triangles)
            T = .Triangles(i)
            p1 = .points(T.iP1)
            p2 = .points(T.iP2)
            p3 = .points(T.iP3)
            aPB.Line (tx + p1.X * sc, ty + p1.Y * sc)-(tx + p2.X * sc, ty + p2.Y * sc)
            aPB.Line -(tx + p3.X * sc, ty + p3.Y * sc)
            aPB.Line -(tx + p1.X * sc, ty + p1.Y * sc)
            c = c + 3
        Next
        
    End With
    'Debug.Print c / 2
End Function
Public Function DrawPoint3_xy(aPB As PictureBox, p As Point3, color As Long)
    If sc = 0 Then sc = 37
    Dim tx As Double: tx = aPB.ScaleWidth / 2
    Dim ty As Double: ty = aPB.ScaleHeight / 2
    aPB.ForeColor = color
    aPB.Circle (tx + p.X * sc, ty + p.Y * sc), 3
End Function

Public Function DrawObj3D_xz(aPB As PictureBox, obj As Object3D, color As Long)

    Dim i As Long
    Dim T As Triangle
    Dim p1 As Point3, p2 As Point3, p3 As Point3
    'Dim sc As Double: sc = 37 * 5
    If sc = 0 Then sc = 37
    Dim tx As Double: tx = aPB.ScaleWidth / 2
    Dim tz As Double: tz = aPB.ScaleHeight / 2
    aPB.ForeColor = color
    With obj
    
        For i = LBound(.Triangles) To UBound(.Triangles)
            T = .Triangles(i)
            p1 = .points(T.iP1)
            p2 = .points(T.iP2)
            p3 = .points(T.iP3)
            aPB.Line (tx + p1.X * sc, tz + -p1.Z * sc)-(tx + p2.X * sc, tz + -p2.Z * sc)
            aPB.Line -(tx + p3.X * sc, tz + -p3.Z * sc)
            aPB.Line -(tx + p1.X * sc, tz + -p1.Z * sc)
        Next
        
    End With
    
End Function
Public Function DrawPoint3_xz(aPB As PictureBox, p As Point3, color As Long)
    If sc = 0 Then sc = 37
    Dim tx As Double: tx = aPB.ScaleWidth / 2
    Dim tz As Double: tz = aPB.ScaleHeight / 2
    aPB.ForeColor = color
    aPB.Circle (tx + p.X * sc, tz + p.Z * sc), 3
End Function

Public Sub ShadeObj3D_projected(aPB As PictureBox, aProj As Matrix34, obj As Object3D, color As Long)
    'hier code zum zeichnen einfügen
    Dim T As Triangle
    Dim n As Long: n = (UBound(obj.Triangles) - LBound(obj.Triangles) + 1) * 3
    ReDim p(0 To n - 1) As Point2L
    Dim i As Long, c As Long
    For i = LBound(obj.Triangles) To UBound(obj.Triangles)
        T = obj.Triangles(i)
        p(c) = Point3_Projection(obj.points(T.iP1), aProj): c = c + 1
        p(c) = Point3_Projection(obj.points(T.iP2), aProj): c = c + 1
        p(c) = Point3_Projection(obj.points(T.iP3), aProj): c = c + 1
    Next
    Dim rv As Long
    ReDim colors(0 To 256) As Long

    aPB.ScaleMode = vbPixels
    aPB.FillStyle = vbSolid
    aPB.FillColor = vbRed 'color
    Dim hDC As Long: hDC = aPB.hDC
    For i = 0 To UBound(p) Step 3
        'if obj.Triangles(i).iNorm
        rv = Polygon(hDC, p(i), 3)
    Next

End Sub

Public Sub DrawObj3D_projected(aPB As PictureBox, aProj As Matrix34, obj As Object3D, color As Long)
    'hier code zum zeichnen einfügen
    Dim T As Triangle
    Dim n As Long: n = (UBound(obj.Triangles) - LBound(obj.Triangles) + 1) * 3
    ReDim p(0 To n - 1) As Point2L
    Dim i As Long, c As Long
    For i = LBound(obj.Triangles) To UBound(obj.Triangles)
        T = obj.Triangles(i)
        p(c) = Point3_Projection(obj.points(T.iP1), aProj): c = c + 1
        p(c) = Point3_Projection(obj.points(T.iP2), aProj): c = c + 1
        p(c) = Point3_Projection(obj.points(T.iP3), aProj): c = c + 1
    Next
    Dim rv As Long
    Dim hDC As Long: hDC = aPB.hDC
    For i = 0 To UBound(p) Step 3
        rv = Polyline(hDC, p(i), 3)
    Next
    ''wie schauts aus mit den Flächennormalen?
End Sub
Public Function DrawPoint3_projected(aPB As PictureBox, aProj As Matrix34, pt As Point3, color As Long)
    If sc = 0 Then sc = 37
    Dim tx As Double: tx = aPB.ScaleWidth / 2
    Dim ty As Double: ty = aPB.ScaleHeight / 2
    aPB.ForeColor = color
    Dim p As Point2L: p = Point3_Projection(pt, aProj)
    aPB.Circle (p.X, p.Y), 3
End Function

Public Sub ZoomIn(Optional ByVal zoomfact As Double = 1.5)
    sc = sc * zoomfact
End Sub
Public Sub ZoomOut(Optional ByVal zoomfact As Double = 1.5)
    sc = sc / zoomfact
End Sub
'Public Function DrawObj3D_yz(aPB As PictureBox, obj As Object3D)
'    Dim i As Long
'    Dim t As Triangle
'    Dim p1 As Point3, p2 As Point3, p3 As Point3
'    Dim sc As Double: sc = 37
'    Dim ty As Double: ty = aPB.ScaleWidth / 2
'    Dim tz As Double: tz = aPB.ScaleHeight / 2
'    With obj
'        For i = LBound(.Triangles) To UBound(.Triangles)
'            t = .Triangles(i)
'            p1 = .Points(t.iP1)
'            p2 = .Points(t.iP2)
'            p3 = .Points(t.iP3)
'            aPB.Line (ty + p1.y * sc, tz + p1.Z * sc)-(ty + p2.y * sc, tz + p2.Z * sc)
'            aPB.Line -(ty + p3.y * sc, tz + p3.Z * sc)
'            aPB.Line -(ty + p1.y * sc, tz + p1.Z * sc)
'        Next
'    End With
'End Function

'Parsers
Public Function ParsePoints(VArr) As Point3()
    Dim i As Long
    ReDim p(LBound(VArr) To UBound(VArr)) As Point3
    For i = LBound(VArr) To UBound(VArr)
        p(i) = New_Point3_va(VArr(i))
    Next
    ParsePoints = p
End Function
Public Function ParseTriangles(VArr) As Triangle()
    Dim i As Long
    ReDim T(LBound(VArr) To UBound(VArr)) As Triangle
    For i = LBound(VArr) To UBound(VArr)
        T(i) = New_Triangle_va(VArr(i))
    Next
    ParseTriangles = T
End Function

'Public Function Distance_o3d(obj As Object3D) As Double
'    Dim p As Point3
'    With obj
'        p = .Points(0)
'    End With
'End Function
Public Function Distance(p As Point3) As Double
    Distance = Sqr(p.X * p.X + p.Y * p.Y + p.Z * p.Z)
End Function

Public Function Rotate_xy(points() As Point3, ByVal x0 As Double, ByVal y0 As Double, ByVal alpha_rad As Double) As Point3() 'Object3D
    Dim i As Long
    ReDim p(LBound(points) To UBound(points)) As Point3
    Dim sin_a As Double: sin_a = Sin(alpha_rad)
    Dim cos_a As Double: cos_a = Cos(alpha_rad)
    Dim dx As Double, dy As Double
    For i = LBound(p) To UBound(p)
        dx = points(i).X - x0
        dy = points(i).Y - y0
        p(i).X = x0 + dx * cos_a - dy * sin_a
        p(i).Y = y0 + dx * sin_a + dy * cos_a
        p(i).Z = points(i).Z
    Next
    Rotate_xy = p
End Function

