VERSION 5.00
Begin VB.UserControl LabelAngle 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2355
   ScaleHeight     =   1215
   ScaleWidth      =   2355
   ToolboxBitmap   =   "UClbangle.ctx":0000
End
Attribute VB_Name = "LabelAngle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#####################################################
'  LABEL ANGLE v 1.0
'  by LITO  (c)2003
'  http://eureka.ya.com/lito2002web
'  Limits: only between 0-90º
'#####################################################

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
' lfFaceName(32) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFaceName As String * 33
End Type

'Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal Height As Long, ByVal Width As Long, ByVal Escapement As Long, ByVal Orientation As Long, ByVal Weight As Long, ByVal Italic As Long, ByVal Underline As Long, ByVal StrikeOut As Long, ByVal CharSet As Long, ByVal OutputPrecision As Long, ByVal ClipPrecision As Long, ByVal Quality As Long, ByVal PitchAndFamily As Long, ByVal Face As String) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const FW_BOLD = 700    '''''
Private Const FW_NORMAL = 400  '''''


''para traducción de colores
'Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
'Private Const CLR_INVALID = &HFFFF


Public Enum CaptionBack
    [Transparent] = 0
    [Opaque] = 1
End Enum
Public Enum LosBordes
    [None]
    [3D]
End Enum
Public Enum LosColoStyle
    [Independent] = 0
    [AmbientBack] = 1
    [AmbientTotal] = 2
End Enum


'Default Property Values:
Const Pi As Double = 3.14159265358979
Const m_def_Angle = 90
Const m_def_AutoSize As Boolean = True
Const m_def_ColorStyle As Integer = 2
Const m_def_BackColor As Long = &H8000000F
Const m_def_ForeColor As Long = 0
Const m_def_Enabled As Boolean = True
Const m_def_Font = "Arial"
Const m_def_BackStyle As Integer = 0 'transparent
Const m_def_BorderStyle As Integer = 0  'none
Const m_def_PosX As Long = 100
Const m_def_PosY As Long = 1000

'Property Variables:
Dim m_Angle As Single
Dim m_Caption As String
Dim m_AutoSize As Boolean
Dim m_ColorStyle As LosColoStyle
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As CaptionBack
Dim m_BorderStyle As LosBordes
Dim m_PosX As Long
Dim m_PosY As Long

'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse encima de un objeto."
Event DblClick()
Attribute DblClick.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse y después lo vuelve a presionar y liberar sobre un objeto."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Ocurre cuando el usuario presiona una tecla mientras un objeto tiene el enfoque."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Ocurre cuando el usuario presiona y libera una tecla ANSI."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Ocurre cuando el usuario libera una tecla mientras un objeto tiene el enfoque."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Ocurre cuando el usuario presiona el botón del mouse mientras un objeto tiene el enfoque."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Ocurre cuando el usuario mueve el mouse."
'Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)



'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
  m_Angle = m_def_Angle
  m_Caption = Extender.Name
  m_AutoSize = m_def_AutoSize
  m_BackColor = m_def_BackColor
  m_ForeColor = m_def_ForeColor
  m_Enabled = m_def_Enabled
  m_BackStyle = m_def_BackStyle
  m_BorderStyle = m_def_BorderStyle
  m_PosX = m_def_PosX
  m_PosY = m_def_PosY
   
  Set m_Font = New StdFont
  m_Font.Name = m_def_Font
  Set UserControl.Font = m_Font

End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
  m_Angle = .ReadProperty("Angle", m_def_Angle)
  m_Caption = .ReadProperty("Caption", Extender.Name)
  m_AutoSize = .ReadProperty("Autosize", m_def_AutoSize)
  m_ColorStyle = .ReadProperty("ColorStyle", m_def_ColorStyle)
  m_BackColor = .ReadProperty("BackColor", m_def_BackColor)
  m_ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
  m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
  Set m_Font = .ReadProperty("Font", "Arial")
  m_BackStyle = .ReadProperty("BackStyle", m_def_BackStyle)
  m_BorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)
  m_PosX = .ReadProperty("PosX", m_def_PosX)
  m_PosY = .ReadProperty("PosY", m_def_PosY)
End With
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  .WriteProperty "Angle", m_Angle, m_def_Angle
  .WriteProperty "Caption", m_Caption, Extender.Name
  .WriteProperty "AutoSize", m_AutoSize, m_def_AutoSize
  .WriteProperty "ColorStyle", m_ColorStyle, m_def_ColorStyle
  .WriteProperty "BackColor", m_BackColor, m_def_BackColor
  .WriteProperty "Caption", m_Caption, Extender.Name
  .WriteProperty "ForeColor", m_ForeColor, m_def_ForeColor
  .WriteProperty "Enabled", m_Enabled, m_def_Enabled
  .WriteProperty "Font", m_Font, Ambient.Font
  .WriteProperty "BackStyle", m_BackStyle, m_def_BackStyle
  .WriteProperty "BorderStyle", m_BorderStyle, m_def_BorderStyle
  .WriteProperty "PosX", m_PosX, m_def_PosX
  .WriteProperty "PosY", m_PosY, m_def_PosY
End With
End Sub







'MemberInfo=12,0,0,0
Public Property Get Angle() As Single
  Angle = m_Angle
End Property
Public Property Let Angle(ByVal New_Angle As Single)
  m_Angle = New_Angle
  PropertyChanged "Angle"
  DrawControl
End Property

'MemberInfo=13,0,0,0
Public Property Get Caption() As String
  Caption = m_Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
  m_Caption = New_Caption
  PropertyChanged "Caption"
  DrawControl
End Property


Public Property Get AutoSize() As Boolean
  AutoSize = m_AutoSize
End Property
Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
  m_AutoSize = New_AutoSize
  PropertyChanged "AutoSize"
  DrawControl
End Property

Public Property Get ColorStyle() As LosColoStyle
  ColorStyle = m_ColorStyle
End Property
Public Property Let ColorStyle(ByVal New_ColorStyle As LosColoStyle)
  m_ColorStyle = New_ColorStyle
  PropertyChanged "ColorStyle"
  DrawControl
End Property

'MemberInfo=8,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
  BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  m_BackColor = New_BackColor
  PropertyChanged "BackColor"
  UserControl.BackColor = m_BackColor
  DrawControl
End Property

'MemberInfo=8,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
  ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  m_ForeColor = New_ForeColor
  PropertyChanged "ForeColor"
  UserControl.ForeColor = m_ForeColor
  DrawControl
End Property

'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  m_Enabled = New_Enabled
  PropertyChanged "Enabled"
End Property

'MemberInfo=6,0,0,0
Public Property Get Font() As Font
  Set Font = m_Font
End Property
Public Property Set Font(ByVal New_Font As Font)
  Set m_Font = New_Font
  PropertyChanged "Font"
  DrawControl
End Property

'MemberInfo=7,0,0,0
Public Property Get BackStyle() As CaptionBack
Attribute BackStyle.VB_Description = "Indica si un control Label o el color de fondo de un control Shape es transparente u opaco."
  BackStyle = m_BackStyle
End Property
Public Property Let BackStyle(ByVal New_BackStyle As CaptionBack)
  m_BackStyle = New_BackStyle
  PropertyChanged "BackStyle"
  DrawControl
End Property

'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As LosBordes
Attribute BorderStyle.VB_Description = "Devuelve o establece el estilo del borde de un objeto."
  BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As LosBordes)
  m_BorderStyle = New_BorderStyle
  PropertyChanged "BorderStyle"
  DrawControl
End Property

'MemberInfo=12,0,0,0
Public Property Get PosX() As Long
Attribute PosX.VB_Description = "Devuelve o establece las coordenadas horizontales para el siguiente método Print o Draw."
  PosX = m_PosX
End Property
Public Property Let PosX(ByVal New_PosX As Long)
  m_PosX = New_PosX
  PropertyChanged "PosX"
  UserControl.CurrentX = m_PosX
  DrawControl
End Property

'MemberInfo=12,0,0,0
Public Property Get PosY() As Single
Attribute PosY.VB_Description = "Devuelve o establece las coordenadas verticales para el siguiente método Print o Draw."
  PosY = m_PosY
End Property
Public Property Let PosY(ByVal New_PosY As Single)
  m_PosY = New_PosY
  PropertyChanged "PosY"
  UserControl.CurrentY = m_PosY
  DrawControl
End Property


'MemberInfo=5
Public Sub Cls()
   UserControl.Cls
End Sub


Private Sub UserControl_Resize()
  DrawControl
End Sub


Private Sub UserControl_AmbientChanged(PropertyName As String)
If m_ColorStyle <> Independent Or m_BackStyle = Transparent Then 'if parent forms backcolor changes
    'tambien tenemos que actualizar en este control la zona entre el título y el cuerpo
    If PropertyName = "BackColor" Or PropertyName = "ForeColor" Then
       DrawControl
    End If
End If
End Sub


Private Sub DrawControl()
  On Error GoTo GetOut
  If m_Font.Name = "" Then UserControl_InitProperties
  
  'Colores-----------------------------
  Dim DrColor1 As Long, DrColor2 As Long, DrColor3 As Long
  ' Fore, Back, Antialias
  UserControl.BorderStyle = m_BorderStyle
  UserControl.BackStyle = 1  'opaque
  DrColor1 = m_ForeColor
  DrColor2 = m_BackColor

  If m_ColorStyle = AmbientBack Then
    DrColor2 = Ambient.BackColor
  End If
  If m_ColorStyle = AmbientTotal Then
    DrColor1 = Ambient.ForeColor
    DrColor2 = Ambient.BackColor
  End If
  UserControl.ForeColor = DrColor1
  UserControl.BackColor = DrColor2
  If BackStyle = Transparent Then
    DrColor2 = Ambient.BackColor
    'ya que si el texto coincide con el color del fondo, tambien se vuelve transparente
    UserControl.BackColor = Abs(m_ForeColor - 3) 'garantizamos de que no coincida
  End If
'  'If AntiAlias Then
'  If 1 = 2 Then
'    DrColor3 = ColorPromedia(DrColor1, DrColor2)
'  End If
  UserControl.Cls
  '------------------------------------
  
  'para autosize-----------------------
  With UserControl.Font
    .Name = m_Font.Name
    .Size = m_Font.Size
    .Bold = m_Font.Bold
    .Italic = m_Font.Italic
  End With
  Dim txtAlto As Long, txtAncho As Long
  Dim UCalto As Long, UCancho As Long
  Dim Co As Single, Se As Single, Bord As Single
  If m_Caption = "" Then 'if a caption no exists
    txtAlto = 1
    txtAncho = 1
  Else
    txtAlto = UserControl.TextHeight(m_Caption)
    txtAncho = UserControl.TextWidth(m_Caption)
  End If
  'el punto de inserción del texto es en la esquina sup.izq.
  Dim Ang As Single, Sg As Integer
  Sg = Sgn(m_Angle)
  Ang = (m_Angle + Sg * 180) Mod 360 - (Sg * 180)
  Co = Cos(Ang / 180 * Pi)
  Se = Sin(Ang / 180 * Pi)
  If m_AutoSize = True Then
    If Ang >= 0 Then
      m_PosX = 0 'txtAlto  * Se
      m_PosY = Abs(txtAncho * Se)
      If Abs(Ang) > 90 Then
        m_PosX = Abs(txtAncho * Co)
        m_PosY = m_PosY + Abs(txtAlto * Co)
      End If
    Else
      m_PosX = -(txtAlto * Se)
      m_PosY = 0
      If Abs(Ang) > 90 Then
        m_PosX = m_PosX + Abs(txtAncho * Co)
        m_PosY = Abs(txtAlto * Co)
      End If
    End If
    Bord = 0
    If m_BorderStyle = [3D] Then Bord = Screen.TwipsPerPixelX * 2 '2 pixels de borde
    UCancho = Abs(txtAlto * Se) + Abs(txtAncho * Co) + Bord * 2
    UCalto = Abs(txtAncho * Se) + Abs(txtAlto * Co) + Bord * 2
    UserControl.Size (UCancho), (UCalto)
  End If
  '------------------------------------
  
  
  'Datos de la fuente------------------
  Dim F As LOGFONT, hPrevFont As Long, hFont As Long ', FontName As String
  F.lfEscapement = 10 * Ang 'rotation angle, in tenths
  F.lfFaceName = m_Font.Name + Chr$(0)  'null terminated
  F.lfHeight = (m_Font.Size * -20) / Screen.TwipsPerPixelY
  If m_Font.Bold Then F.lfWeight = FW_BOLD Else F.lfWeight = FW_NORMAL
  F.lfItalic = m_Font.Italic
  F.lfStrikeOut = m_Font.Strikethrough
  F.lfUnderline = m_Font.Underline
  hFont = CreateFontIndirect(F)
      If hFont = 0 Then Exit Sub
  hPrevFont = SelectObject(UserControl.hdc, hFont)
  '------------------------------------
  
'  'Anti-alias--------------------------
'If 1 = 2 Then
'  Dim i As Integer
'  i = Screen.TwipsPerPixelX
' With UserControl
'  .ForeColor = DrColor3
'  .CurrentX = m_PosX + i: .CurrentY = m_PosY: Print m_Caption
'  .CurrentX = m_PosX - i: .CurrentY = m_PosY: Print m_Caption
'  .CurrentX = m_PosX: .CurrentY = m_PosY + i: Print m_Caption
'  .CurrentX = m_PosX: .CurrentY = m_PosY - i: Print m_Caption
'  .CurrentX = m_PosX + i: .CurrentY = m_PosY + i: Print m_Caption
'  .CurrentX = m_PosX - i: .CurrentY = m_PosY + i: Print m_Caption
'  .CurrentX = m_PosX - i: .CurrentY = m_PosY - i: Print m_Caption
'  .CurrentX = m_PosX + i: .CurrentY = m_PosY - i: Print m_Caption
' End With
'End If
'  '------------------------------------
  
  'Dibujo de la fuente-----------------
' hFont = GetClientRect(UserControl.hwnd, rc)
' hFont = TextOut(UserControl.hdc, m_PosX, m_PosX, (m_Caption), Len(m_Caption))
  UserControl.ForeColor = DrColor1
  UserControl.CurrentX = m_PosX
  UserControl.CurrentY = m_PosY
  Print m_Caption
  '------------------------------------
  
  'Clean up, restore original font
  hFont = SelectObject(UserControl.hdc, hPrevFont)
  DeleteObject hFont
  
  If m_BackStyle = Transparent Then
    UserControl.MaskColor = UserControl.BackColor
    UserControl.MaskPicture = UserControl.Image
    UserControl.BackStyle = 0
  End If
  
  Exit Sub
GetOut:
  If Err.Number = 91 Then Resume Next

End Sub



'Private Function ColorPromedia(ByVal Color1 As Long, ByVal Color2 As Long) As Long
''para Promediar un color entre otros 2.
'Dim Red As Long, Blue As Long, Green As Long, A$
'Dim Red1 As Long, Blue1 As Long, Green1 As Long
'Dim Red2 As Long, Blue2 As Long, Green2 As Long
'
'If Len(Hex$(Color1)) > 6 Then 'si es un color del sistema
'  Color1 = TranslateColour(Color1)
'End If
'If Len(Hex$(Color2)) > 6 Then 'si es un color del sistema
'  Color2 = TranslateColour(Color2)
'End If
'
'Blue1 = ((Color1 \ &H10000) Mod &H100)
'Green1 = ((Color1 \ &H100) Mod &H100)
'Red1 = (Color1 And &HFF)
'Blue2 = ((Color2 \ &H10000) Mod &H100)
'Green2 = ((Color2 \ &H100) Mod &H100)
'Red2 = (Color2 And &HFF)
'
'Blue = (Blue1 + Blue2) / 2
'Green = (Green1 + Green2) / 2
'Red = (Red1 + Red2) / 2
'
'ColorPromedia = Red + 256& * Green + 65536 * Blue
'
'End Function
'
'
''---------------------------------------------------------------------------------------
'' Procedure : TranslateColour
'' DateTime  : 12/10/2003
'' Author    : Drew (aka The Bad One)
'' Purpose   : Used to convert Automation colours to a Windows (long) colour.
''---------------------------------------------------------------------------------------
'Private Function TranslateColour(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
'   If TranslateColor(oClr, hPal, TranslateColour) Then
'       TranslateColour = CLR_INVALID
'   End If
'End Function
'

