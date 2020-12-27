Attribute VB_Name = "M2D"
Option Explicit

Public Type Point2
    X As Double
    Y As Double
End Type

Public Type Point2L
    X As Long
    Y As Long
End Type

Public Pi
Public Pi2

Public Const eps As Double = 0.000000000001

Public Function New_Point2(ByVal X As Double, ByVal Y As Double) As Point2
    
    New_Point2.X = X: New_Point2.Y = Y
    
End Function
    
Public Function New_RndPoint2() As Point2
    
    New_RndPoint2.X = Rnd * 2 - 1: New_RndPoint2.Y = Rnd * 2 - 1
    
End Function
    
Public Function Trapez_Fläche(ByVal a0 As Double, ByVal a1 As Double, ByVal b As Double) As Double
    'berechnet die Fläche eines stehenden Trapezes mit ungleichen aber parallelen Seiten a0 und a1
    Trapez_Fläche = b * (a0 + a1) / 2
End Function

Public Function Trapez_Schwerpunkt_y(ByVal a0 As Double, ByVal a1 As Double) As Double
    'berechnet die y-Koordinate des Schwerpunkts eines Trapezes mit parallelen Seiten a0 und a1
    'vertauschen?
    If a0 > a1 Then: Dim T As Double: T = a1: a1 = a0: a0 = T
    Trapez_Schwerpunkt_y = (a1 ^ 2 - a0 ^ 2 + a0 * (a1 + 2 * a0)) / (3 * (a1 + a0))
End Function

Public Function Trapez_Schwerpunkt_x(ByVal a0 As Double, ByVal a1 As Double, ByVal b As Double) As Double
    'berechnet die y-Koordinate des Schwerpunkts eines Trapezes mit parallelen Seiten a0 und a1
    'vertauschen?
    If a0 > a1 Then: Dim T As Double: T = a1: a1 = a0: a0 = T
    'B6/3 * (B5+2*B3)/(B5+B3)
    Trapez_Schwerpunkt_x = b / 3 * (a1 + 2 * a0) / (a1 + a0)
End Function

Public Function Trapez_Schwerpunkt(ByVal a0 As Double, ByVal a1 As Double, ByVal b As Double, out_A As Double) As Point2 'Double
    'berechnet die Koordinaten des Schwerpunkts eines Trapezes mit parallelen Seiten a0 und a1, und der Breite b
    'a0,a1-vertauschen?
    If a0 > a1 Then: Dim T As Double: T = a1: a1 = a0: a0 = T
    'B6/3 * (B5+2*B3)/(B5+B3)
    Dim a1_a0 As Double: a1_a0 = a1 + a0
    
    Trapez_Schwerpunkt.X = b / 3 * (a1 + 2 * a0) / a1_a0
    Trapez_Schwerpunkt.Y = (a1 ^ 2 - a0 ^ 2 + a0 * (a1 + 2 * a0)) / (3 * a1_a0)
    
    out_A = b * a1_a0 / 2
    
End Function

Public Function Ellipse_Fläche(ByVal A As Double, ByVal b As Double) As Double
    'berechnet die Fläche einer über die Halbachsen definierten Ellipse
    
    Ellipse_Fläche = A * b * Pi
    
End Function

Public Function HalbEllipse_Schwerpunkt(ByVal A As Double) As Double
    'berechnet die y-Koordinate des Schwerpunkts einer über der x-Achse liegenden Halbellipse
    'a ist die Halbachse in y-Richtung
    HalbEllipse_Schwerpunkt = 4 * A / (3 * Pi)
    
    'diese Formel gilt auch für die x-Koordinate des Schwerpunkts wenn für a die andere Halbachse angegeben wird
    'diese Formel gilt auch für die y-Koordinate des Schwerpunkts eines Kreises
End Function

Public Function CreateRechteck(ByVal a_x As Double, ByVal b_y As Double) As Point2()
    
    ReDim Rechteck(0 To 4) As Point2
    
    Dim i As Long
    Rechteck(i).X = a_x / 2:  Rechteck(i).Y = -b_y / 2: i = i + 1
    Rechteck(i).X = a_x / 2:  Rechteck(i).Y = b_y / 2:  i = i + 1
    Rechteck(i).X = -a_x / 2: Rechteck(i).Y = b_y / 2:  i = i + 1
    Rechteck(i).X = -a_x / 2: Rechteck(i).Y = -b_y / 2: i = i + 1
    Rechteck(i).X = a_x / 2:  Rechteck(i).Y = -b_y / 2
    
    CreateRechteck = Rechteck
    
End Function

Public Function CreateEllipsePolygon(ByVal A As Double, ByVal b As Double, _
                                     ByVal vonAlpha As Integer, ByVal bisAlpha As Integer) As Point2()
    'erzeugt ein Polygon für einen Teil einer Ellipse in ganze Grad-Schritte, z.B. von 0°-180°
    Dim n As Integer: n = bisAlpha - vonAlpha
    'umdrehen ?
    If n < 0 Then: Dim T As Integer: T = vonAlpha: vonAlpha = bisAlpha: bisAlpha = T: n = -n
    ReDim p(0 To n) As Point2
    
    Dim alpha_rad As Double
    Dim alpha As Integer: alpha = vonAlpha
    Dim i As Long
    For i = 0 To n
        alpha_rad = (alpha + i) * Pi / 180
        p(i).X = A * Cos(alpha_rad)
        p(i).Y = b * Sin(alpha_rad)
    Next
    
    CreateEllipsePolygon = p
    
End Function

Public Function CreateHalbEi(ByVal a_li As Double, ByVal b As Double, ByVal a_re As Double) As Point2()
    'erzeugt ein Polygon für ein Halb-Ei über der x-Achse
    
    ReDim p(0 To 180) As Point2
    Dim i As Long, a_Rad As Double
    
    For i = 0 To 90 '180
        a_Rad = i * Pi / 180
        p(i).X = a_re * Cos(a_Rad)
        p(i).Y = b * Sin(a_Rad)
    Next
    
    For i = 91 To 180
        a_Rad = i * Pi / 180
        p(i).X = a_li * Cos(a_Rad)
        p(i).Y = b * Sin(a_Rad)
    Next
    
    CreateHalbEi = p
    
End Function

Public Function CreateEi(ByVal r_x As Double, ByVal r_yo As Double, ByVal r_yu As Double) As Point2()
    
    If Pi = 0 Then Pi = 4 * Atn(1)
    
    Dim a_Rad As Double
    Dim i As Long
    ReDim Ei(0 To 360) As Point2
    
    For i = 0 To 180
        a_Rad = i * Pi / 180
        Ei(i).X = r_x * Cos(a_Rad)
        Ei(i).Y = r_yo * Sin(a_Rad)
    Next
    
    For i = 181 To 360
        a_Rad = i * Pi / 180
        Ei(i).X = r_x * Cos(a_Rad)
        Ei(i).Y = r_yu * Sin(a_Rad)
    Next
    
    CreateEi = Ei
    
End Function
Public Function CreateEiEi(ByVal r_xl As Double, ByVal r_xr As Double, ByVal r_yo As Double, ByVal r_yu As Double) As Point2()
    
    If Pi = 0 Then Pi = 4 * Atn(1)
    
    Dim a_Rad As Double
    Dim i As Long
    ReDim Ei(0 To 360) As Point2
    
    For i = 0 To 90
        a_Rad = i * Pi / 180
        Ei(i).X = r_xr * Cos(a_Rad)
        Ei(i).Y = r_yo * Sin(a_Rad)
    Next
    For i = 91 To 180
        a_Rad = i * Pi / 180
        Ei(i).X = r_xl * Cos(a_Rad)
        Ei(i).Y = r_yo * Sin(a_Rad)
    Next
    
    For i = 181 To 270
        a_Rad = i * Pi / 180
        Ei(i).X = r_xl * Cos(a_Rad)
        Ei(i).Y = r_yu * Sin(a_Rad)
    Next
    For i = 271 To 360
        a_Rad = i * Pi / 180
        Ei(i).X = r_xr * Cos(a_Rad)
        Ei(i).Y = r_yu * Sin(a_Rad)
    Next
    
    CreateEiEi = Ei
    
End Function

Public Function CreateRndPolygon(ByVal nEck As Long) As Point2()
    
    Dim i As Long
    ReDim p(0 To nEck) As Point2
    
    For i = 0 To nEck - 1
        p(i) = New_RndPoint2
    Next
    
    p(nEck) = p(0)
    CreateRndPolygon = p
    
End Function

Public Function Verschieben(poly() As Point2, ByVal X As Double, ByVal Y As Double) As Point2()
    'verschiebt ein polygon
    Dim i As Long
    ReDim p(LBound(poly) To UBound(poly)) As Point2
    For i = LBound(poly) To UBound(poly)
        p(i).X = poly(i).X + X
        p(i).Y = poly(i).Y + Y
    Next
    Verschieben = p
End Function

'Public Function Rotate_xy(points() As Point2, ByVal x0 As Double, ByVal y0 As Double, ByVal alpha_rad As Double) As Point2()
'
'End Function
Public Function Rotieren(poly() As Point2, ByVal x0 As Double, ByVal y0 As Double, ByVal alpha_rad As Double) As Point2()
    'rotiert ein Polygon poly um einen Punkt x0,y0, um den Winkel alpha gegeben in radians
    Dim i As Long
    ReDim p(LBound(poly) To UBound(poly)) As Point2
    Dim sin_a As Double: sin_a = Sin(alpha_rad)
    Dim cos_a As Double: cos_a = Cos(alpha_rad)
    Dim dx As Double, dy As Double
    For i = LBound(poly) To UBound(poly)
        dx = poly(i).X - x0
        dy = poly(i).Y - y0
        p(i).X = x0 + dx * cos_a - dy * sin_a
        p(i).Y = y0 + dx * sin_a + dy * cos_a
    Next
    Rotieren = p
End Function

Public Function Skalieren(poly() As Point2, ByVal scalefaktor As Double) As Point2()
    Dim i As Long
    ReDim p(LBound(poly) To UBound(poly)) As Point2
    For i = LBound(poly) To UBound(poly)
        p(i).X = poly(i).X * scalefaktor
        p(i).Y = poly(i).Y * scalefaktor
    Next
    Skalieren = p
End Function

'hier wären jetzt noch schön so Sachen wie
'* Punkt_in_Polygon
'* Polygon verschneiden
'* Polygon vereinigen

Public Function Schwerpunkt(fx() As Point2, out_A As Double) As Point2
'berechnet den Schwerpunkt und die Fläche eines Polygons in x,y-Koordinaten
    Dim x0 As Double, y0 As Double
    Dim x1 As Double, y1 As Double
    Dim dx As Double, dy As Double
    Dim dhx As Double, dhy As Double
    Dim dAx As Double, dAy As Double
    Dim xsi As Double, ysi As Double
    Dim Sum_dAx As Double, Sum_dAy As Double
    Dim Sum_xsi_dAx As Double, Sum_ysi_dAy As Double
    Dim l As Long: l = LBound(fx)
    Dim u As Long: u = UBound(fx)
    If u - l = 1 Then Exit Function
    x0 = fx(l).X
    y0 = fx(l).Y
    Dim i As Long
    For i = l + 1 To u
        'für xs:               | 'für ys:
        y1 = fx(i).Y:            x1 = fx(i).X
        dy = (y1 - y0):          dx = (x1 - x0)
        dhx = (x1 + x0) / 2:     dhy = (y1 + y0) / 2
        dAx = dy * dhx:          dAy = dx * dhy
        Sum_dAx = Sum_dAx + dAx: Sum_dAy = Sum_dAy + dAy
        
        'Schwerpunkt Trapez: xs = (a^2 - b^2 + b*(a + 2*b))/(3*(a + b))
        If (x1 + x0) = 0 Then
            xsi = 0
        Else
            xsi = (x1 * x1 - x0 * x0 + x0 * (x1 + 2 * x0)) / (3 * (x1 + x0))
        End If
        
        If y1 + y0 = 0 Then
            ysi = 0
        Else
            ysi = (y1 * y1 - y0 * y0 + y0 * (y1 + 2 * y0)) / (3 * (y1 + y0))
        End If
        
        Sum_xsi_dAx = Sum_xsi_dAx + xsi * dAx
        Sum_ysi_dAy = Sum_ysi_dAy + ysi * dAy

        'Punkt übergeben
        x0 = x1
        y0 = y1
    Next
    With Schwerpunkt
        .X = Sum_xsi_dAx / Sum_dAx
        If Abs(.X) < eps Then .X = 0
        .Y = Sum_ysi_dAy / Sum_dAy
        If Abs(.Y) < eps Then .Y = 0
    End With
    out_A = (Abs(Sum_dAx) + Abs(Sum_dAy)) / 2
End Function



