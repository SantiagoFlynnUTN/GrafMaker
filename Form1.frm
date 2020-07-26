VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Mercurio Magic Studio Spliter."
   ClientHeight    =   11445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   763
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1395
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command12 
      Caption         =   "Recargar Datos"
      Height          =   495
      Left            =   1680
      TabIndex        =   44
      Top             =   9840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command11 
      Height          =   375
      Index           =   4
      Left            =   5520
      TabIndex        =   43
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Oeste"
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   42
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Este"
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   41
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Sur"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   40
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Norte"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   39
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Misc"
      Height          =   375
      Index           =   5
      Left            =   6840
      TabIndex        =   38
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Muerte"
      Height          =   375
      Index           =   4
      Left            =   5520
      TabIndex        =   37
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Caminata"
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   36
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Reposo"
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   35
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Golpeado"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   34
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Ataque"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   33
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1920
      TabIndex        =   32
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   5160
      TabIndex        =   28
      Top             =   8280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Deseleccionar Todos"
      Height          =   360
      Left            =   7800
      TabIndex        =   27
      Top             =   1320
      Width           =   1650
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Ver cuadricula"
      Height          =   360
      Left            =   7800
      TabIndex        =   26
      Top             =   840
      Width           =   1650
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<"
      Height          =   720
      Left            =   18120
      TabIndex        =   25
      Top             =   2280
      Width           =   690
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">"
      Height          =   720
      Left            =   19080
      TabIndex        =   24
      Top             =   2280
      Width           =   690
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   3615
      Left            =   240
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1325
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3120
      Width           =   19935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Generar"
      Height          =   600
      Left            =   7080
      TabIndex        =   22
      Top             =   9720
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   8040
      TabIndex        =   20
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5400
      TabIndex        =   18
      Top             =   9240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5400
      TabIndex        =   17
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Text            =   "256"
      Top             =   9240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Text            =   "256"
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nuevo Grafico"
      Height          =   360
      Left            =   240
      TabIndex        =   11
      Top             =   8280
      Width           =   2130
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Split Frame"
      Height          =   360
      Left            =   5520
      TabIndex        =   10
      Top             =   1320
      Width           =   1890
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calcular"
      Height          =   360
      Left            =   5520
      TabIndex        =   8
      Top             =   840
      Width           =   1890
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H0080FF80&
      ForeColor       =   &H0000FFFF&
      Height          =   480
      Index           =   0
      Left            =   360
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cargar Gráfico"
      CausesValidation=   0   'False
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label11 
      Height          =   1815
      Left            =   4440
      TabIndex        =   45
      Top             =   9840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2160
      TabIndex        =   31
      Top             =   2760
      Width           =   45
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frames seleccionados:"
      Height          =   195
      Left            =   360
      TabIndex        =   30
      Top             =   2760
      Width           =   1620
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   195
      Left            =   4200
      TabIndex        =   29
      Top             =   8400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Frames:"
      Height          =   195
      Left            =   6480
      TabIndex        =   21
      Top             =   9000
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filas:"
      Height          =   195
      Left            =   4440
      TabIndex        =   19
      Top             =   9360
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Columnas:"
      Height          =   195
      Left            =   4440
      TabIndex        =   16
      Top             =   9000
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alto de lienzo:"
      Height          =   195
      Left            =   1680
      TabIndex        =   15
      Top             =   9360
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ancho de lienzo:"
      Height          =   195
      Left            =   1680
      TabIndex        =   14
      Top             =   9000
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   525
      Left            =   345
      Top             =   2145
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de frames:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1380
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grafico alto:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grafico ancho:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frame alto:"
      Height          =   195
      Left            =   3000
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frame ancho:"
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   960
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

CommonDialog1.InitDir = App.PATH & "\Graficos_raw\"
CommonDialog1.Filter = "Apps (*.bmp)|*.bmp|All files (*.*)|*.*"
CommonDialog1.DefaultExt = "txt"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen

Picture2.Cls
If Len(CommonDialog1.FileName) > 0 Then
Abrir CommonDialog1.FileName
End If

End Sub

Private Sub Command10_Click()
Dim i As Long
FramesSeleccionados = 0

For i = 1 To NumFrames

    FrameSelectos(i) = False
    Picture1(i).BackColor = vbWhite
    
    
Next i
Label14.Caption = FramesSeleccionados

End Sub

Private Sub Command11_Click(Index As Integer)
    DireccionAnim = Command11(Index).Caption
    newGrafico
End Sub

Private Sub Command12_Click()
 newGrafico
End Sub

Private Sub Command2_Click()

If Val(Text11.Text) > 0 Then
    NumFrames = Val(Text11.Text)
    Frameswidth = Info.biWidth / NumFrames
    FramesHeight = Info.biHeight
    Text1.Text = Frameswidth
    Text2.Text = FramesHeight
ElseIf Val(Text1.Text) > 0 And Val(Text2.Text) > 0 Then
    Frameswidth = Text1.Text
    FramesHeight = Text2.Text
    NumFrames = Info.biWidth / Frameswidth
    Text11.Text = NumFrames

Exit Sub
End If

Calcular
        FramesSeleccionados = NumFrames
        Form1.Label14.Caption = FramesSeleccionados
        
End Sub

Private Sub Command3_Click()
If selectedframe > 0 And selectedframe <= NumFrames Then
    ObtenerFrame selectedframe
End If
End Sub

Private Sub newGrafico()
If Frameswidth = 0 Or FramesHeight = 0 Then
    Command2_Click
    Exit Sub
End If


Text5.Text = Val(Text3.Text / Frameswidth)
Text6.Text = Val(Text4.Text / FramesHeight)
Text10.Text = TipoAnim & DireccionAnim & "-" & FramesSeleccionados & "-" & Frameswidth & "-" & FramesHeight & "-" & Text5.Text & "-" & Text6.Text
Text7.Text = FramesSeleccionados

Label11.Caption = "Frames por medidas: " & vbCrLf & "128: " & Val(Val(128 / Frameswidth) * Val(128 / FramesHeight)) & vbCrLf & "256: " & Val(Val(256 / Frameswidth) * Val(256 / FramesHeight)) & vbCrLf & "512: " & Val(Val(512 / Frameswidth) * Val(512 / FramesHeight))

End Sub

Private Sub Command4_Click()

newGrafico

Command5.Visible = True
Command12.Visible = True

Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True

Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text10.Visible = True
End Sub

Private Sub Command5_Click()
Dim i As Long
Dim c As Long
nWidth = Text3.Text
nHeight = Text4.Text
nColumnas = Val(Text5.Text)
nFilas = Val(Text6.Text)
newnFrames = Val(Text7.Text)

If newnFrames > nFilas * nColumnas Then
    MsgBox "El numero de frames supera el grafico."
    Exit Sub
End If

If nColumnas * Frameswidth > nWidth Then
    MsgBox "El ancho del grafico es incorrecto."
    Exit Sub
End If

If nFilas * FramesHeight > nHeight Then
    MsgBox "El alto del grafico es incorrecto."
    Exit Sub
End If
ReDim FRAME(1 To NumFrames)
For i = 1 To NumFrames

    If FrameSelectos(i) Then
        c = c + 1
        FRAME(c) = i
    End If
    

Next i
If c > newnFrames Then
    MsgBox "El numero de frames selectos sobrepasa los frames del nuevo grafico."
    Exit Sub
End If
Module1.CrearGrafico

End Sub

Private Sub Command6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If pic_off_x <= Info.biWidth - Frameswidth * 2 Then
pic_off_x = pic_off_x + Frameswidth
If pic_off_x > Info.biWidth - Picture2.Width Then
    'pic_off_x = Info.biWidth - Picture2.Width
End If
Form1.Picture2.Cls
Form1.Picture2.PaintPicture PICT, 0, 0, , , pic_off_x
        If bVerCuadricula Then DibujarCuadricula
End If
End Sub

Private Sub Command7_Click()
pic_off_x = pic_off_x - Frameswidth
If pic_off_x <= 0 Then pic_off_x = 0
Form1.Picture2.PaintPicture PICT, 0, 0, , , pic_off_x
        If bVerCuadricula Then DibujarCuadricula
End Sub



Private Sub Command8_Click(Index As Integer)
    TipoAnim = Command8(Index).Caption
    newGrafico
End Sub

Private Sub Command9_Click()

bVerCuadricula = Not bVerCuadricula

DibujarCuadricula


End Sub


Private Sub Form_Click()
    selectedframe = 0
    Shape1.Visible = False
    
End Sub



Private Sub Picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
        If Index >= 1 Then
    If FrameSelectos(Index) Then
        Picture1(Index).BackColor = vbWhite
        FrameSelectos(Index) = False
        FramesSeleccionados = FramesSeleccionados - 1
    Else
        FrameSelectos(Index) = True
        Picture1(Index).BackColor = vbGreen
        FramesSeleccionados = FramesSeleccionados + 1
    End If
    End If
    Label14.Caption = FramesSeleccionados

ElseIf Button = vbLeftButton Then

    If ((Index - 1) * Frameswidth) < pic_off_x Then
        pic_off_x = (Index - 1) * Frameswidth
        Form1.Picture2.PaintPicture PICT, 0, 0, , , pic_off_x
    ElseIf ((Index - 1) * Frameswidth) > pic_off_x + Form1.Picture2.Width Then
        pic_off_x = (Index - 1) * Frameswidth
        If pic_off_x > Info.biWidth - Picture2.Width Then pic_off_x = Info.biWidth - Picture2.Width
        Form1.Picture2.PaintPicture PICT, 0, 0, , , pic_off_x
    End If
    
    selectedframe = Index
    Shape1.Left = Picture1(Index).Left - 1
    Shape1.Top = Picture1(Index).Top - 1
    Shape1.Visible = True
    If bVerCuadricula Then DibujarCuadricula
    
    

End If

End Sub

Private Sub Text11_Change()
 Command2_Click
End Sub
