VERSION 5.00
Begin VB.UserControl ctlToolBox 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3525
   ControlContainer=   -1  'True
   ScaleHeight     =   2955
   ScaleWidth      =   3525
   ToolboxBitmap   =   "ctlToolBox.ctx":0000
   Begin VB.Image imgToggle 
      Height          =   180
      Left            =   30
      Picture         =   "ctlToolBox.ctx":0312
      Top             =   60
      Width           =   150
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   2610
      Picture         =   "ctlToolBox.ctx":0671
      Top             =   30
      Width           =   240
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "jghghjkghjk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   210
      TabIndex        =   0
      Top             =   0
      Width           =   2370
   End
   Begin VB.Shape shpTitleBarBack 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "ctlToolBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mdblOrigSize As Double
Dim mblnToggle As Boolean
Dim mstrTitle As String
Dim mintTitleAlign As Integer
Dim mlngTitleBarBackColor As Long
Dim mlngTitleBoxBackColor As Long
Dim mstrToggleImage As String
Dim mstrIcon As String

Public Event Expand()
Public Event Collapse()
Public Event Resize()

Private Sub imgToggle_Click()
    If mblnToggle Then
        UserControl.Height = mdblOrigSize
        mblnToggle = False
        RaiseEvent Expand
    Else
        UserControl.Height = lblTitle.Height + 60
        mblnToggle = True
        RaiseEvent Collapse
    End If
End Sub

Private Sub UserControl_Initialize()
    mstrTitle = ""
    mintTitleAlign = 1
    mdblOrigSize = UserControl.Height
    mlngTitleBarBackColor = &H80000005
    mlngTitleBoxBackColor = &H808080
    mstrToggleImage = App.Path & "\toggle.gif"
    mstrIcon = App.Path & "\icon.ico"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mstrTitle = PropBag.ReadProperty("Title", "")
    mintTitleAlign = CInt(PropBag.ReadProperty("TitleAlign", "1"))
    mdblOrigSize = CDbl(PropBag.ReadProperty("ExpandHeight", UserControl.Height))
    mlngTitleBarBackColor = CLng(PropBag.ReadProperty("TitleBarBackColor", &H80000005))
    mlngTitleBoxBackColor = CLng(PropBag.ReadProperty("TitleBoxBackColor", &H808080))
    mstrToggleImage = PropBag.ReadProperty("ToggleImage", App.Path & "\toggle.gif")
    mstrIcon = PropBag.ReadProperty("Icon", App.Path & "\icon.ico")
    
    UserControl.Height = mdblOrigSize
    lblTitle = mstrTitle
    lblTitle.Alignment = mintTitleAlign
    lblTitle.BackColor = mlngTitleBoxBackColor
    shpTitleBarBack.BackColor = mlngTitleBarBackColor
    If Dir(mstrToggleImage) <> "" Then imgToggle.Picture = LoadPicture(mstrToggleImage)
    If Dir(mstrIcon) <> "" Then imgIcon.Picture = LoadPicture(mstrIcon)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    imgIcon.Left = UserControl.ScaleWidth - imgIcon.Width - 30
    lblTitle.Width = UserControl.ScaleWidth - imgToggle.Width - imgIcon.Width - 120
    RaiseEvent Resize
End Sub

Public Property Get Title() As String
    Title = mstrTitle
End Property

Public Static Property Let Title(strTitle As String)
    mstrTitle = strTitle
    lblTitle = mstrTitle
    PropertyChanged Title
End Property

Public Sub Collapse()
    UserControl.Height = lblTitle.Height + 60
    mblnToggle = True
End Sub

Public Sub Expand()
    UserControl.Height = mdblOrigSize
    mblnToggle = False
End Sub

Public Property Get TitleAlign() As Integer
    TitleAlign = mintTitleAlign
End Property

Public Property Let TitleAlign(ByVal TitleAlignment As Integer)
    If TitleAlignment >= 1 And TitleAlignment <= 3 Then
        mintTitleAlign = TitleAlignment
        lblTitle.Alignment = mintTitleAlign
        PropertyChanged TitleAlign
    End If
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Title", mstrTitle, ""
    PropBag.WriteProperty "TitleAlign", mintTitleAlign
    PropBag.WriteProperty "ExpandHeight", mdblOrigSize
    PropBag.WriteProperty "TitleBarBackColor", shpTitleBarBack.BackColor
    PropBag.WriteProperty "TitleBoxBackColor", lblTitle.BackColor
    PropBag.WriteProperty "ToggleImage", mstrToggleImage
    PropBag.WriteProperty "Icon", mstrIcon
End Sub

Public Property Get ExpandHeight() As Variant
    ExpandHeight = mdblOrigSize
End Property

Public Property Let ExpandHeight(ByVal newHeight As Variant)
    mdblOrigSize = newHeight
End Property

Public Property Get TitleBarBackColor() As Long
    TitleBarBackColor = mlngTitleBarBackColor
End Property

Public Property Let TitleBarBackColor(ByVal NewColor As Long)
    mlngTitleBarBackColor = NewColor
    shpTitleBarBack.BackColor = mlngTitleBarBackColor
    PropertyChanged TitleBarBackColor
End Property

Public Property Get TitleBoxBackColor() As Long
    TitleBoxBackColor = mlngTitleBoxBackColor
End Property

Public Property Let TitleBoxBackColor(ByVal NewColor As Long)
    mlngTitleBoxBackColor = NewColor
    lblTitle.BackColor = mlngTitleBoxBackColor
    PropertyChanged TitleBoxBackColor
End Property

Public Property Get ToggleImage() As String
    ToggleImage = mstrToggleImage
End Property

Public Property Let ToggleImage(ByVal ImagePath As String)
    If Dir(ImagePath) <> "" Then
        mstrToggleImage = ImagePath
        imgToggle.Picture = LoadPicture(ImagePath)
        PropertyChanged ToggleImage
    End If
End Property

Public Property Get Icon() As String
    Icon = mstrIcon
End Property

Public Property Let Icon(ByVal ImagePath As String)
    If Dir(ImagePath) <> "" Then
        mstrIcon = ImagePath
        imgIcon.Picture = LoadPicture(ImagePath)
        PropertyChanged Icon
    End If
End Property
