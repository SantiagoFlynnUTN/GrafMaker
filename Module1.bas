Attribute VB_Name = "Module1"
Option Explicit
Public PICT As StdPicture
Public pic_off_x As Long
Public FramesSeleccionados As Long
Public Type BITMAPFILEHEADER

      bfType As Integer
      bfSize As Long
      bfReserved1 As Integer
      bfReserved2 As Integer
      bfOffBits As Long

End Type
Public NewGRafCount As Long
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                       ByRef source As Any, _
                                       ByVal length As Long)
Public Type BITMAPINFOHEADER

      biSize As Long
      biWidth As Long
      biHeight As Long
      
      biPlanes As Integer
      biBitCount As Integer
      
      biCompression As Long
      biSizeImage As Long
      biXPelsPerMeter As Long
      
      biYPelsPerMeter As Long
      biClrUsed As Long
      biClrImportant As Long

End Type
'40
'4
'14

Public Type RGBQUAD

      rgbBlue As Byte
      rgbGreen As Byte
      rgbRed As Byte
      rgbReserved As Byte

End Type
Public RGBq As RGBQUAD
Public Head As BITMAPFILEHEADER
Public Info As BITMAPINFOHEADER
Public Frameswidth As Long
Public FramesHeight As Long
Public NumFrames As Integer
Public FrameSelectos() As Boolean
Public FF As Integer
Public selectedframe As Integer
Public SplitCount As Long

Public nWidth As Long
Public nHeight As Long
Public newnFrames As Integer
Public nFilas As Integer
Public nColumnas As Integer
Public FRAME() As Integer
Public bVerCuadricula As Boolean




Public Sub Calcular()
Dim i As Long



If NumFrames >= 1 Then
ReDim FrameSelectos(1 To NumFrames)
With Form1
Load .Picture1(1)
.Picture1(1).Left = .Picture1(0).Left
.Picture1(1).Top = .Picture1(0).Top
    .Picture1(1).BackColor = vbGreen
    .Picture1(1).Visible = True
    FrameSelectos(1) = True
If NumFrames >= 2 Then
For i = 2 To NumFrames

    Load .Picture1(i)
    .Picture1(i).Left = .Picture1(i - 1).Left + 34
    .Picture1(i).Top = .Picture1(i - 1).Top
    .Picture1(i).Visible = True
    .Picture1(i).BackColor = vbGreen
    FrameSelectos(i) = True
Next i

End If
End With
End If
End Sub
Public Sub Abrir(ByVal PATH As String)

Set PICT = LoadPicture(PATH)

Form1.Picture2.PaintPicture PICT, 0, 0

FF = FreeFile


    Open PATH For Binary Access Read Write Lock Read Write As #FF
            
        Get FF, , Head
        Get FF, , Info
        Get FF, , RGBq
        
        Form1.Label3.Caption = "Grafico Ancho: " & Info.biWidth
        Form1.Label4.Caption = "Grafico Alto: " & Info.biHeight
        
    

End Sub
Public Sub getFrame(ByVal FRAME As Integer, ByRef Data() As Byte)


Dim i As Long
Dim iStart As Long
Dim bRGB(2) As Byte
Dim FrameArray() As Byte
Dim Fila() As Byte
Dim count As Long
Dim NextStart As Long
Dim asize As Long
Dim fsize As Long
Dim resto As Long
Dim p As Long
Dim fresto As Long
resto = Info.biWidth Mod 4
fresto = Frameswidth Mod 4

fsize = ((Frameswidth * 3)) '+ fresto
asize = ((Frameswidth * 3)) * FramesHeight

ReDim Data(0 To asize - 1) As Byte
iStart = (Head.bfOffBits + 1) + ((3 * Frameswidth) * (FRAME - 1))
ReDim Fila(0 To fsize - 1)
NextStart = iStart

For i = 1 To FramesHeight
    Get FF, NextStart, Fila()

    CopyMemory Data(count), Fila(0), fsize
    count = count + fsize
    NextStart = NextStart + (Info.biWidth * 3) + resto

Next i
End Sub
Public Sub ObtenerFrame(ByVal FRAME As Integer)

Dim i As Long
Dim iStart As Long
Dim bRGB(2) As Byte
Dim FrameArray() As Byte
Dim Fila() As Byte
Dim count As Long
Dim NextStart As Long
Dim asize As Long
Dim fsize As Long
Dim resto As Long
Dim p As Long
Dim fresto As Long
resto = Info.biWidth Mod 4
fresto = Frameswidth Mod 4

fsize = ((Frameswidth * 3)) + fresto
asize = ((Frameswidth * 3) + fresto) * FramesHeight

ReDim FrameArray(0 To asize - 1) As Byte
iStart = (Head.bfOffBits + 1) + ((3 * Frameswidth) * (FRAME - 1))
ReDim Fila(0 To fsize - 1)
NextStart = iStart

For i = 1 To FramesHeight
    Get FF, NextStart, Fila()
    If fresto > 0 Then
    For p = 1 To fresto
    Fila(fsize - 1 - (p - 1)) = 0
    
    Next p
    End If
    CopyMemory FrameArray(count), Fila(0), fsize
    count = count + fsize
    NextStart = NextStart + (Info.biWidth * 3) + resto

Next i

Dim Nuevo As Integer

Nuevo = FreeFile
Dim nhead As BITMAPFILEHEADER
Dim nInfo As BITMAPINFOHEADER
Dim quad As RGBQUAD


nInfo.biHeight = FramesHeight
nInfo.biWidth = Frameswidth
nInfo.biPlanes = 1
nInfo.biBitCount = Info.biBitCount
nInfo.biCompression = Info.biCompression
nInfo.biClrUsed = Info.biClrUsed
nInfo.biClrImportant = Info.biClrImportant
nInfo.biSize = Info.biSize
nInfo.biSizeImage = asize + 54

quad.rgbBlue = RGBq.rgbBlue
quad.rgbGreen = RGBq.rgbGreen
quad.rgbRed = RGBq.rgbRed
quad.rgbReserved = RGBq.rgbReserved

nhead.bfSize = asize + 54
nhead.bfType = Head.bfType
nhead.bfOffBits = 54




Open App.PATH & "\" & SplitCount & ".bmp" For Binary Access Write As #Nuevo

    Put Nuevo, , nhead
    Put Nuevo, , nInfo
    Put Nuevo, , FrameArray()


Close Nuevo
SplitCount = SplitCount + 1
End Sub

Public Sub CrearGrafico()
Dim Nuevo As Integer

Nuevo = FreeFile
Dim nhead As BITMAPFILEHEADER
Dim nInfo As BITMAPINFOHEADER
Dim quad As RGBQUAD
Dim WiDRest As Long
Dim asize As Long
Dim GARRAY() As Byte
Dim k As Long
Dim fsize As Long
Dim frest As Long
Dim count As Long
Dim Pos_En_Fila As Long
WiDRest = nWidth Mod 4


fsize = (Frameswidth * 3) * FramesHeight
asize = ((nWidth * 3) + WiDRest) * nHeight

Dim j As Long

ReDim GARRAY(asize - 1)
Dim Pos_En_Columna As Long
Dim Data() As Byte
ReDim Data(fsize - 1)
Dim Fila As Long
Dim Columna As Long
Dim pxFilaActual As Long
For k = nFilas To 1 Step -1
Dim SrcpxFila As Long
Dim margenInferior As Long
Dim o As Long
margenInferior = (nHeight - (nFilas * FramesHeight)) * ((nWidth * 3) + WiDRest)
For o = (((k - 1) * nColumnas) + 1) To (((k - 1) * nColumnas)) + nColumnas
    If o > newnFrames Then Exit For
    If FRAME(o) = 0 Then Exit For
    getFrame FRAME(o), Data

    Pos_En_Fila = ((Frameswidth * 3) * (o - (((k - 1) * nColumnas) + 1)))
    Pos_En_Columna = (FramesHeight * ((nWidth * 3) + WiDRest) * (nFilas - k)) + margenInferior

    For j = 0 To FramesHeight - 1
        pxFilaActual = (j * ((nWidth * 3) + WiDRest))
        SrcpxFila = (Frameswidth * 3) * j
        CopyMemory GARRAY(Pos_En_Columna + pxFilaActual + Pos_En_Fila), Data(SrcpxFila), Frameswidth * 3
    Next j

Next o
Next k


nInfo.biHeight = nHeight 'FramesHeight 'nHeight
nInfo.biWidth = nWidth 'Frameswidth 'nWidth
nInfo.biPlanes = 1
nInfo.biBitCount = Info.biBitCount
nInfo.biCompression = Info.biCompression
nInfo.biClrUsed = Info.biClrUsed
nInfo.biClrImportant = Info.biClrImportant
nInfo.biSize = Info.biSize
nInfo.biSizeImage = asize + 54 '(Frameswidth + Frameswidth Mod 4) * FramesHeight + 54 'asize + 54

quad.rgbBlue = RGBq.rgbBlue
quad.rgbGreen = RGBq.rgbGreen
quad.rgbRed = RGBq.rgbRed
quad.rgbReserved = RGBq.rgbReserved

nhead.bfSize = asize + 54 '(Frameswidth + Frameswidth Mod 4) * FramesHeight + 54 'asize + 54
nhead.bfType = Head.bfType
nhead.bfOffBits = 54


Nuevo = FreeFile

Open App.PATH & "\" & Form1.Text10.Text & ".bmp" For Binary Access Write As #Nuevo

    Put Nuevo, , nhead
    Put Nuevo, , nInfo
    Put Nuevo, , GARRAY()
    'Put Nuevo, , Data

Close #Nuevo
NewGRafCount = NewGRafCount + 1


End Sub
Public Sub DibujarCuadricula()
Dim i As Long
Dim color As Long
Dim ActualFrame As Long
Dim FrameInicial As Long
Form1.Picture2.DrawWidth = 2
If Frameswidth <= 0 Then Exit Sub

FrameInicial = pic_off_x / Frameswidth

For i = 1 To Form1.Picture2.Width / Frameswidth
    ActualFrame = FrameInicial + i
    If ActualFrame = selectedframe Or ActualFrame = selectedframe - 1 Then
        color = vbBlue
    Else
        color = vbWhite
    End If
    'j = Form1.Picture2.Point((i * Frameswidth), k)
    Form1.Picture2.Line (i * Frameswidth, 0)-(i * Frameswidth, FramesHeight), color

Next i

End Sub
