VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18555
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   18555
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox Check4 
      Caption         =   "Shade"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   120
      Width           =   855
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Perspective"
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "RotSpeed -"
      Height          =   375
      Left            =   15720
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "RotSpeed +"
      Height          =   375
      Left            =   14520
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Rotate"
      Height          =   255
      Left            =   13440
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   480
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Eggify"
      Height          =   375
      Left            =   12120
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show all inner shapes"
      Height          =   255
      Left            =   8880
      TabIndex        =   8
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Subdivide -"
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Subdivide +"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnZoomOut 
      Caption         =   "Zoom Out"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnZoomIn 
      Caption         =   "Zoom In"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Euler?"
      Height          =   375
      Left            =   10920
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000005&
      ForeColor       =   &H00000000&
      Height          =   7935
      Left            =   8760
      ScaleHeight     =   7875
      ScaleWidth      =   8715
      TabIndex        =   4
      Top             =   480
      Width           =   8775
   End
   Begin VB.OptionButton Option2 
      Caption         =   "xz"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "xy"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   7935
      Left            =   120
      ScaleHeight     =   7875
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   480
      Width           =   8655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   17040
      TabIndex        =   15
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim icosphere1 As Object3D
Dim icosph2()  As Object3D
Dim eggics2()  As Object3D
Dim colors()  As Long
Dim c As Long
Dim egg_x_fact_pos As Double
Dim egg_x_fact_neg As Double
Dim egg_z_fact_pos As Double
Dim egg_z_fact_neg As Double
Dim rotangle As Double
Dim rotspeed As Double
Dim m_Projection As Matrix34
Dim bIsFullscreenPb1 As Boolean
Dim bIsFullscreenPb2 As Boolean
Dim Center As Point3
Dim alpha_x As Double

Private Sub Form_Load()
    egg_x_fact_pos = 1
    egg_x_fact_neg = 1
    egg_z_fact_pos = 1
    egg_z_fact_neg = 1
    icosphere1 = CreateIcosahedron(IcosahedronPoint_r(1))
    Dim i As Long
    ReDim icosph2(0 To 6)
    ReDim eggics2(0 To 6)
    ReDim colors(0 To 6)
    icosph2(0) = icosphere1 'kopieren
    'Debug.Print "i: " & i & " | " & CheckEulerPolyeder(newics(i))
    For i = 1 To 6
        icosph2(i) = Icosahedron_subdivide(icosph2(i - 1))
        'Debug.Print "i: " & i & " | p: " & CheckEulerPolyeder(newics(i))
        colors(i) = RGB(Rnd * 256, Rnd * 256, Rnd * 256)
    Next
    eggics2 = icosph2 'kopieren
    M2D.Pi = 4 * Atn(1)
    M2D.Pi2 = 4 * Atn(1) * 2
    alpha_x = Pi / 2
    
    rotspeed = 1
    Timer1.Interval = 10
    Timer1.Enabled = False
    Option1.Caption = "xy"
    Option2.Caption = "xz"
    Check3.Caption = "Perspective"
    'Option3.Caption = "Projection"
    BtnZoomIn.Caption = "Zoom In"
    BtnZoomOut.Caption = "Zoom Out"
    Command1.Caption = "Subdivide +"
    Command2.Caption = "Subdivide -"
    Command3.Caption = "Euler?"
    Picture1.ScaleMode = vbPixels
    Picture1.AutoRedraw = True
    Picture2.ScaleMode = vbPixels
    Picture2.AutoRedraw = True
    UpdateLblRot
    Draw
'    Dim tx As Double: tx = Picture2.ScaleWidth / 2
'    Dim tz As Double: tz = Picture2.ScaleHeight / 2
'    m_Projection = New_Matrix34Camera(3000, 1, 1, tx, tz, 0, 0, 15, 0, 0, 0)
'    Draw1
'    Draw2
    
End Sub
Private Sub Option1_Click()
    Draw
End Sub
Private Sub Option2_Click()
    Draw
End Sub
Private Sub Check3_Click()
    Draw
End Sub
Private Sub Check4_Click()
    Draw2
End Sub
Private Sub Check2_Click()
    Timer1.Enabled = Check2.Value = vbChecked
    Draw2
End Sub

Private Sub UpdateLblRot()
    Label1.Caption = "rs: " & Round(rotspeed, 2)
End Sub
Private Sub BtnZoomIn_Click()
    M3D.ZoomIn
    Draw
End Sub
Private Sub BtnZoomOut_Click()
    M3D.ZoomOut
    Draw
End Sub
Private Sub Command1_Click()
    c = IIf(c = 6, 0, c + 1)
    Picture1.ForeColor = colors(c)
    Draw2
End Sub
Private Sub Command2_Click()
    c = IIf(c = 0, 6, c - 1)
    Picture1.ForeColor = colors(c)
    Draw2
End Sub
Private Sub Check1_Click()
    Draw
End Sub
Private Sub Command3_Click()
    MsgBox M3D.CheckEulerPolyeder(icosph2(c))
End Sub
Private Sub Command4_Click()
    Dim s As String: s = InputBox("Type in: x/z; fact_pos; fact_neg", "Eggify a polyeder", "x; 1.618; 1.221")
    If Len(s) = 0 Then Exit Sub
    Dim sa() As String: sa = Split(s, ";")
    Dim xz   As String: xz = Trim(sa(0))
    Dim fpos As Double: fpos = Val(Trim(sa(1)))
    Dim fneg As Double: fneg = Val(Trim(sa(2)))
    Dim i As Long
    If xz = "x" Then
        egg_x_fact_pos = fpos
        egg_x_fact_neg = fneg
    ElseIf xz = "z" Then
        egg_z_fact_pos = fpos
        egg_z_fact_neg = fneg
    End If
    For i = 0 To 6
        eggics2(i) = Eggify_x(Eggify_z(icosph2(i), 0, egg_z_fact_pos, egg_z_fact_neg), 0, egg_x_fact_pos, egg_x_fact_neg)
    Next
    'so jetzt noch den Center berechnen
    Dim EiEi() As Point2: EiEi = CreateEiEi(egg_x_fact_pos, egg_x_fact_neg, egg_z_fact_pos, egg_z_fact_neg)
    Dim A As Double
    Dim cnt As Point2: cnt = Schwerpunkt(EiEi, A)
    Center.X = -cnt.X
    'Center.Y
    Center.Z = -cnt.Y
    Draw2
End Sub
Private Sub Command5_Click()
    rotspeed = rotspeed + IIf(-1 <= rotspeed And rotspeed <= 0.9, 0.1, 1)
    UpdateLblRot
End Sub
Private Sub Command6_Click()
    rotspeed = rotspeed - IIf(-0.9 <= rotspeed And rotspeed <= 1, 0.1, 1)
    UpdateLblRot
End Sub

Private Sub Form_Resize()
    Dim brdr: brdr = 8 * Screen.TwipsPerPixelX
    Dim l: l = brdr
    Dim T: T = Picture1.Top
    Dim W: W = (Me.ScaleWidth - 2 * l) / 2 'IIf(bIsFullscreenPb1 Or bIsFullscreenPb2, 1, 2)
    Dim H: H = Me.ScaleHeight - T - brdr
    If W > 0 And H > 0 Then
        'W = W * IIf(bIsFullscreenPb1, 2, IIf(bIsFullscreenPb2, 0, 1))
        If bIsFullscreenPb1 Then
            Picture1.Move l, T, 2 * W, H
            Picture1.ZOrder 0
        Else
            Picture1.Move l, T, W, H
            If Not bIsFullscreenPb2 Then l = l + W
        End If
        InitProjection Picture1
        Draw1
    End If
    If W > 0 And H > 0 Then
        'W = W * IIf(bIsFullscreenPb2, 2, 1)
        If bIsFullscreenPb2 Then
            Picture2.Move l, T, W * 2, H
            Picture2.ZOrder 0
        Else
            Picture2.Move l, T, W, H
        End If
'            Command1.Left = l
'            Command2.Left = l + Command1.Width
'            Check1.Left = l + Command1.Width + Command2.Width + brdr
'            Command3.Left = Check1.Left + Check1.Width + brdr
'            Command4.Left = Command3.Left + Command3.Width
'            BtnZoomOut.Left = l - BtnZoomOut.Width
'            BtnZoomIn.Left = l - BtnZoomOut.Width - BtnZoomIn.Width
        InitProjection Picture2
        Draw2
        'End If
    End If
End Sub
Private Sub InitProjection(aPB As PictureBox)
    Dim tx As Double: tx = aPB.ScaleWidth / 2
    Dim tz As Double: tz = aPB.ScaleHeight / 2
    Dim ax As Double
    If Option1.Value Then
        ax = 0
    ElseIf Option2.Value Then
        ax = alpha_x 'Pi / 2
    End If
    m_Projection = New_Matrix34Camera(300, 60 / M3D.sc, 60 / M3D.sc, tx, tz, 0, 0, -5, ax, 0, 0)
End Sub
Private Sub Draw()
    If Check3.Value = vbChecked Then InitProjection Picture1
    Draw1
    If Check3.Value = vbChecked Then InitProjection Picture2
    Draw2
End Sub
Private Sub Draw1()
    Picture1.Cls
    If Check3.Value = vbChecked Then
        M3D.DrawObj3D_projected Picture1, m_Projection, icosphere1, vbBlack
    Else
        If Option1.Value Then
            M3D.DrawObj3D_xy Picture1, icosphere1, vbBlack
        ElseIf Option2.Value Then
            M3D.DrawObj3D_xz Picture1, icosphere1, vbBlack
        End If
    End If
End Sub

Private Sub Draw2()
    Picture2.Cls
    Dim i As Long
    Dim start As Long: start = IIf(Check1.Value = vbChecked, 0, c)
    Dim rotobj As Object3D
    Dim color As Long: color = colors(c)
    
    If Check3.Value = vbChecked Then
        M3D.DrawPoint3_projected Picture2, m_Projection, Center, color
    Else
        If Option1.Value Then
            M3D.DrawPoint3_xy Picture2, Center, color
        Else
            M3D.DrawPoint3_xz Picture2, Center, color
        End If
    End If
    For i = start To c
        rotobj = eggics2(i)
        rotobj.points = M3D.Rotate_xy(rotobj.points, Center.X, Center.Y, rotangle * Pi / 180)
        color = colors(i)
        If Check3.Value = vbChecked Then
            If Check4.Value = vbChecked Then
                M3D.ShadeObj3D_projected Picture2, m_Projection, rotobj, color
            Else
                M3D.DrawObj3D_projected Picture2, m_Projection, rotobj, color
            End If
        Else
            If Option1.Value Then
                M3D.DrawObj3D_xy Picture2, rotobj, color
            ElseIf Option2.Value Then
                M3D.DrawObj3D_xz Picture2, rotobj, color
            End If
        End If
    Next
End Sub

Private Sub Picture1_DblClick()
    bIsFullscreenPb1 = Not bIsFullscreenPb1
    Form_Resize
End Sub
Private Sub Picture2_DblClick()
    bIsFullscreenPb2 = Not bIsFullscreenPb2
    Form_Resize
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Picture2_KeyDown KeyCode, Shift
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp:   alpha_x = alpha_x + 1 * Pi / 180
    Case vbKeyDown: alpha_x = alpha_x - 1 * Pi / 180
    End Select
    InitProjection Picture2
    Draw
End Sub

Private Sub Timer1_Timer()
    If Check2.Value = vbChecked Then
        rotangle = rotangle + rotspeed
        Draw2
    End If
End Sub
