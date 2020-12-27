Attribute VB_Name = "MCamera"
Option Explicit
'm_Projection = New_TMatrix34Camera(100, 1, 1, 0, 0, 20, 20, 20, New_AngleDeg(0), New_AngleDeg(0), New_AngleDeg(0))

'Public Function New_TMatrix34Camera(ByVal w_b As Double, _
'                                    ByVal s_x As Double, ByVal s_y As Double, _
'                                    ByVal B_x0 As Double, ByVal B_y0 As Double, _
'                                    ByVal O_cx As Double, ByVal O_cy As Double, ByVal O_cz As Double, _
'                                    ByVal alp_X As Double, ByVal bet_Y As Double, ByVal gam_Z As Double) As TMatrix34
'    Dim extr As TMatrix44: extr = New_TMatrix44RotXYZ(alp_X, bet_Y, gam_Z)
'    With extr:        .ad = O_cx:        .bd = O_cy:        .cd = O_cz:    End With
'    'Debug.Print Matrix44_ToStr(extr)
'    Dim intr As TMatrix34: intr = New_TMatrix34(-w_b / s_x, 0, B_x0, 0, _
'                                                0, -w_b / s_y, B_y0, 0, _
'                                                0, 0, 1, 0)
'    New_TMatrix34Camera = Mat34Multiply44(intr, extr)
'End Function
Private m_Name As String
Private m_wb As Double        'focal distance = Brennweite
Private m_AspectRatio1 As Boolean 'True: Seitenverhältnis = 1; False: beliebig
Private m_ScX As Double       'Scalierfaktor X
Private m_ScY As Double       'Scalierfaktor Y
Private m_B0X As Double       '
Private m_B0Y As Double       '
Private m_OcX As Double       'Augpunkt X
Private m_OcY As Double       'Augpunkt Y
Private m_OcZ As Double       'Augpunkt Z
Private m_alpX As Double      'Drehwinkel um X
Private m_betY As Double      'Drehwinkel um Y
Private m_gamZ As Double      'Drehwinkel um Z

Private m_Projection As Matrix34 'die berechnete Projektionsmatrix
'Private m_Pi                  'einfach nur Pi

'Private Sub Class_Initialize()
'    m_Pi = CDec(4 * Atn(1))
''    m_Name = "StdCam": m_wb = 3000: s_x = 1: s_y = 1: O_cz = 15
''    Call InitProjection
'    m_AspectRatio1 = True
'End Sub
'
'Friend Sub New_(ByVal aName As String, ByVal w_b As Double, _
'                ByVal s_x As Double, ByVal s_y As Double, _
'                ByVal B_x0 As Double, ByVal B_y0 As Double, _
'                ByVal O_cx As Double, ByVal O_cy As Double, ByVal O_cz As Double, _
'                ByVal alp_X As Double, ByVal bet_Y As Double, ByVal gam_Z As Double)
'     Name = aName:    m_wb = w_b
'     m_ScX = s_x:     m_ScY = s_y
'     m_B0X = B_x0:    m_B0Y = B_y0
'     m_OcX = O_cx:    m_OcY = O_cy:    m_OcZ = O_cz
'     m_alpX = CheckAngle(alp_X):  m_betY = CheckAngle(bet_Y):  m_gamZ = CheckAngle(gam_Z)
'     Call InitProjection
'End Sub
'Public Sub NewC(ByVal other As Camera)
'     With other
'        Name = other.Name: m_wb = .wb
'        m_ScX = .ScX:    m_ScY = .ScY
'        m_B0X = .B0X:    m_B0Y = .B0Y
'        m_OcX = .OcX:    m_OcY = .OcY:    m_OcZ = .OcZ
'        m_alpX = CheckAngle(.AlpX):  m_betY = CheckAngle(.BetY):  m_gamZ = CheckAngle(.GamZ)
'     End With
'     Call InitProjection
'End Sub
Public Sub New_(ByVal aName As String, ByVal w_b As Double, _
                ByVal s_x As Double, ByVal s_y As Double, _
                ByVal B_x0 As Double, ByVal B_y0 As Double, _
                ByVal O_cx As Double, ByVal O_cy As Double, ByVal O_cz As Double, _
                ByVal alp_X As Double, ByVal bet_Y As Double, ByVal gam_Z As Double)
     m_Name = aName:  m_wb = w_b
     m_ScX = s_x:     m_ScY = s_y
     m_B0X = B_x0:    m_B0Y = B_y0
     m_OcX = O_cx:    m_OcY = O_cy:    m_OcZ = O_cz
     m_alpX = CheckAngle(alp_X):  m_betY = CheckAngle(bet_Y):  m_gamZ = CheckAngle(gam_Z)
     Call InitProjection
End Sub

Private Sub InitProjection()
    m_Projection = New_TMatrix34Camera(m_wb, m_ScX, m_ScY, m_B0X, m_B0Y, m_OcX, m_OcY, m_OcZ, m_alpX, m_betY, m_gamZ)
End Sub

'Public Property Let Name(ByVal Value As String)
'    m_Name = Value
'End Property
'Public Property Get Name() As String
'    Name = m_Name
'End Property
'Public Function Compare(other As Camera) As Long
'    Compare = StrComp(m_Name, other.Name, VbCompareMethod.vbTextCompare)
'End Function
'focal distance = Brennweite
Public Property Let wb(Value As Double)
    m_wb = Value
    InitProjection
End Property
Public Property Get wb() As Double
    wb = m_wb
End Property

Public Property Let AspectRatio1(ByVal Value As Boolean)
    m_AspectRatio1 = Value
    If m_AspectRatio1 Then m_ScY = m_ScX
End Property
Public Property Get AspectRatio1() As Boolean
    AspectRatio1 = m_AspectRatio1
End Property

'scale factor x
Public Property Let ScX(Value As Double)
    m_ScX = Value
    InitProjection
End Property
Public Property Get ScX() As Double
    ScX = m_ScX
End Property

Public Property Let ScY(Value As Double)
    If m_AspectRatio1 Then
        m_ScX = Value
    Else
        m_ScY = Value
    End If
End Property
Public Property Get ScY() As Double
    ScY = IIf(m_AspectRatio1, m_ScX, m_ScY)
End Property




Public Property Let B0X(Value As Double)
    m_B0X = Value
    InitProjection
End Property
Public Property Get B0X() As Double
    B0X = m_B0X
End Property

Public Property Let B0Y(Value As Double)
    m_B0Y = Value
    InitProjection
End Property
Public Property Get B0Y() As Double
    B0Y = m_B0Y
End Property



'Augpunkt
Public Property Let OcX(Value As Double)
    m_OcX = Value
    InitProjection
End Property
Public Property Get OcX() As Double
    OcX = m_OcX
End Property

Public Property Let OcY(Value As Double)
    m_OcY = Value
    InitProjection
End Property
Public Property Get OcY() As Double
    OcY = m_OcY
End Property

Public Property Let OcZ(ByVal Value As Double)
    m_OcZ = Value
    InitProjection
End Property
Public Property Get OcZ() As Double
    OcZ = m_OcZ
End Property


'Drehwinkel x
Public Property Let AlpX(ByVal Value As Double)
    m_alpX = CheckAngle(Value)
    InitProjection
End Property
Public Property Get AlpX() As Double
    AlpX = m_alpX
End Property

'Drehwinkel y
Public Property Let BetY(ByVal Value As Double)
    m_betY = CheckAngle(Value)
    InitProjection
End Property
Public Property Get BetY() As Double
    BetY = m_betY
End Property

'Drehwinkel z
Public Property Let GamZ(ByVal Value As Double)
    m_gamZ = CheckAngle(Value)
    InitProjection
End Property
Public Property Get GamZ() As Double
    GamZ = m_gamZ
End Property


'die Projektionsmatrix
'Friend Property Get Projection() As TMatrix34
'    Projection = m_Projection
'End Property


Public Function ToStr() As String
    ToStr = Name & " " & MMat.Matrix34_ToStr(m_Projection)
End Function


