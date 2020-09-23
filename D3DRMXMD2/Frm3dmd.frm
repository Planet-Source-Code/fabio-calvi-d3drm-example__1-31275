VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Character Animation-By MartWare-FPS:"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   663
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMain 
      Interval        =   1000
      Left            =   1800
      Top             =   1560
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'An example about static characters with DirectX7 D3DRM and X files
'(the characters are taken from Quake 2 models)
'
'Keys to use:
'Up arrow    -> move forward
'Down arrow  -> move back
'Left arrow  -> rotate left
'Right arrow -> rotate right
'PageUp      -> look up
'PageDown    -> look down
'Q           -> rotate chars towards right
'W           -> rotate chars towards left
'Esc         -> quit

Option Explicit
Const pi = 3.1415927
' direct x objects
Dim dx As New DirectX7
Dim dd As DirectDraw4
Dim clip As DirectDrawClipper
Dim d3drm As Direct3DRM3
Dim scene As Direct3DRMFrame3
Dim cam As Direct3DRMFrame3
Dim dev As Direct3DRMDevice3
Dim view As Direct3DRMViewport2
Dim mesh As Direct3DRMMeshBuilder3
Dim XFileTex As Direct3DRMTexture3
Dim LightSpot As Direct3DRMLight
Dim Light As Direct3DRMFrame3
Dim m_objectFrame(5) As Direct3DRMFrame3
Dim m_meshBuilder(5) As Direct3DRMMeshBuilder3
Dim m_object As Direct3DRMMeshBuilder3

Dim Grados As Integer
Dim grado2 As Integer
Dim valore As Integer
Dim corx As Single
Dim corz As Single

' dll calls
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

' Main sub
Private Sub Form_Load()
Dim i As Integer
Dim lngTimer As Long
Dim lngFrameTimer As Long
Dim FPSCounter As Integer
Dim FPS As Integer
Dim lngElapsed As Long

    ShowCursor 0

    ' init directdraw and clipper
    Set dd = dx.DirectDraw4Create("")
    Set clip = dd.CreateClipper(0)
    clip.SetHWnd Me.hWnd
    
    dd.SetDisplayMode 800, 600, 32, 0, DDSDM_DEFAULT
    
    ' init direct d3drm
    Set d3drm = dx.Direct3DRMCreate()
    Set scene = d3drm.CreateFrame(Nothing)
    ' camera
    Set cam = d3drm.CreateFrame(scene)
    corx = 190
    corz = -90
    cam.SetPosition scene, corx, 0, corz
    
    ' add light
    Set Light = d3drm.CreateFrame(scene)
    Light.SetPosition Nothing, 0, 10, 0
    Set LightSpot = d3drm.CreateLightRGB(D3DRMLIGHT_POINT, 1, 0.7, 0.4)
    Light.AddLight LightSpot
    
    ' add a bit of ambient light to the scene
    scene.AddLight d3drm.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.65, 0.65, 0.65)
    '
    ' make viewport and device (3D card is not necessary here)
    Set dev = d3drm.CreateDeviceFromClipper(clip, "IID_IDirect3DRGBDevice", Me.ScaleWidth, Me.ScaleHeight)
    dev.SetQuality D3DRMFILL_SOLID + D3DRMLIGHT_ON + D3DRMSHADE_GOURAUD
    dev.SetTextureQuality D3DRMTEXTURE_LINEAR
    dev.SetDither D_TRUE
    Set view = d3drm.CreateViewport(dev, cam, 0, 0, Me.ScaleWidth, Me.ScaleHeight)
    view.SetBack 2000!
    
    ' create the meshbuilder where add faces(walls) to create the room
    Set mesh = d3drm.CreateMeshBuilder()
    mesh.SetPerspective D_TRUE
    ' add mesh builder to scene
    scene.AddVisual mesh
    
    'create the walls of the room (KOMI method)
    Call MakeWall(d3drm, mesh, -200, -15, 100, 200, -15, 100, 200, -15, -100, -200, -15, -100, "floor", 3, 3, 0, 0, 0) ' grass texture on the floor
    Call MakeWall(d3drm, mesh, 200, 15, 100, -200, 15, 100, -200, 15, -100, 200, 15, -100, "roof", 3, 3, 0, 0, 0)
    Call MakeWall(d3drm, mesh, 200, -15, 100, -200, -15, 100, -200, 15, 100, 200, 15, 100, "wall", 15, 1, 0, 0, 0)
    Call MakeWall(d3drm, mesh, -200, -15, 100, -200, -15, -100, -200, 15, -100, -200, 15, 100, "wall", 15, 1, 0, 0, 0)
    Call MakeWall(d3drm, mesh, -200, -15, -100, 200, -15, -100, 200, 15, -100, -200, 15, -100, "wall", 15, 1, 0, 0, 0)
    Call MakeWall(d3drm, mesh, 200, -15, -100, 200, -15, 100, 200, 15, 100, 200, 15, -100, "wall", 15, 1, 0, 0, 0)

    init_chars
    valore = 10

    Me.Show
    Me.Refresh
    DoEvents
    
    ' start main app loop
    Do While DoEvents()
        
        lngElapsed = dx.TickCount() - lngTimer
        lngTimer = dx.TickCount()
        If dx.TickCount() - lngFrameTimer >= 1000 Then
            lngFrameTimer = dx.TickCount()
            FPS = FPSCounter
            FPSCounter = 0
        Else
            FPSCounter = FPSCounter + 1
        End If
        
        'Don't cross the limits of the room
        '(Well,this is not the best method of collision.
        'I am a beginner with D3DRM and I don't have any book or tutorial about it
        'I am findind a good tutorial about DX7 collision detection (raypick and others);
        'if someone know where is possible to find one and
        'would wish to help me, please email me:FABIOCALVI@YAHOO.COM
        'Thanks! QUID PRO QUOD!)
        If corx < -180 Then corx = -180
        If corx > 180 Then corx = 180
        If corz < -80 Then corz = -80
        If corz > 80 Then corz = 80
       
       'Move forward
        If GetKeyState(vbKeyUp) < -1 Then
           corx = corx + valore * Sin(Grados * pi / 180)
           corz = corz + valore * Cos(Grados * pi / 180)
           cam.SetPosition scene, corx, 0, corz
        End If
        
        'Move back
        If GetKeyState(vbKeyDown) < -1 Then
           corx = corx - valore * Sin(Grados * pi / 180)
           corz = corz - valore * Cos(Grados * pi / 180)
           cam.SetPosition scene, corx, 0, corz
        End If
        
        'Rotate left
        If GetKeyState(vbKeyLeft) < -1 Then
           Grados = Grados - valore
           cam.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -(grado2 * pi / 180)
           cam.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -(valore * pi / 180)
           cam.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, (grado2 * pi / 180)
        End If
        
        'Rotate right
        If GetKeyState(vbKeyRight) < -1 Then
           Grados = Grados + valore
           cam.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -(grado2 * pi / 180)
           cam.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, (valore * pi / 180)
           cam.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, (grado2 * pi / 180)
        End If
        
        'Look up
        If GetKeyState(vbKeyPageUp) < -1 Then
            grado2 = grado2 - valore
            If grado2 > -90 Then
               cam.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -(valore * pi / 180)
            Else
               grado2 = grado2 + valore
            End If
        End If
        
        'Look down
        If GetKeyState(vbKeyPageDown) < -1 Then
            grado2 = grado2 + valore
            If grado2 < 90 Then
               cam.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, (valore * pi / 180)
            Else
               grado2 = grado2 - valore
            End If
        End If
       
        'right(clockwise) rotating character along y axes
        If GetKeyState(vbKeyQ) < -5 Then
           For i = 1 To 5
              m_objectFrame(i).AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, 0.45
           Next
        End If
        'left rotating character along y axes
        If GetKeyState(vbKeyW) < -5 Then
           For i = 1 To 5
              m_objectFrame(i).AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, -0.45
           Next
        End If
        
        ' render the scene
        view.Clear D3DRMCLEAR_ALL
        view.Render scene
        dev.Update
        
        ' check to exit
        If GetKeyState(vbKeyEscape) < -5 Then byebye
  
    Loop
    
End Sub

Private Sub MakeWall(d3drm As Direct3DRM3, mesh As Direct3DRMMeshBuilder3, X1 As Single, Y1 As Single, z1 As Single, x2 As Single, y2 As Single, z2 As Single, x3 As Single, y3 As Single, z3 As Single, x4 As Single, y4 As Single, z4 As Single, TexFile As String, TileX As Single, TileY As Single, r As Single, g As Single, b As Single)
    ' local variables
    Dim f As Direct3DRMFace2
    Dim t As Direct3DRMTexture3
    ' create face
    Set f = d3drm.CreateFace()
    ' add vertexs
    f.AddVertex X1, Y1, z1
    f.AddVertex x2, y2, z2
    f.AddVertex x3, y3, z3
    f.AddVertex x4, y4, z4
    ' get type of file
    If TexFile = "" Then
        ' set colors
        f.SetColorRGB r, g, b
    Else
        ' create texture
        Set t = d3drm.LoadTexture(App.Path & "\" & TexFile & ".bmp")
        ' set u and v values
        f.SetTextureCoordinates 0, 0, 0
        f.SetTextureCoordinates 1, TileX, 0
        f.SetTextureCoordinates 2, TileX, TileY
        f.SetTextureCoordinates 3, 0, TileY
        ' set the texture
        f.SetTexture t
    End If
    ' add face to mesh
    mesh.AddFace f
End Sub
Private Sub init_chars()
Dim i As Integer
    
    'create frame of characters
    For i = 1 To 5
        Set m_objectFrame(i) = d3drm.CreateFrame(scene)
    Next i
    
    'their meshbuilder
    For i = 1 To 5
       Set m_meshBuilder(i) = d3drm.CreateMeshBuilder()
       If i = 1 Then m_meshBuilder(i).LoadFromFile "tre.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
       If i = 2 Or i = 3 Then m_meshBuilder(i).LoadFromFile "weapon.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
       If i = 4 Then m_meshBuilder(i).LoadFromFile "uno.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
       If i = 5 Then m_meshBuilder(i).LoadFromFile "due.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    Next i
    
    'add the meshbuilder to the scene
    For i = 1 To 5
       m_objectFrame(i).AddScale D3DRMCOMBINE_REPLACE, 0.25, 0.25, 0.25
       m_objectFrame(i).AddVisual m_meshBuilder(i)
       m_objectFrame(i).AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, 29.8
    Next i
    
    'set position of characters
    m_objectFrame(1).SetPosition Nothing, 0, -5, -10
    m_objectFrame(2).SetPosition Nothing, 0, -5, -10
    m_objectFrame(1).AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, 45
    m_objectFrame(2).AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, 45
    m_objectFrame(3).SetPosition Nothing, 0, -5, 0
    m_objectFrame(4).SetPosition Nothing, 0, -5, 0
    m_objectFrame(5).SetPosition Nothing, 0, -5, 10
    

End Sub
Private Sub byebye()
    
    Call dd.RestoreDisplayMode
    Call dd.SetCooperativeLevel(frmMain.hWnd, DDSCL_NORMAL)
    
    Set dev = Nothing
    Set d3drm = Nothing
    Set clip = Nothing
    Set dd = Nothing
    Set dx = Nothing

    ShowCursor 1
   
    Unload Me
    End
End Sub
