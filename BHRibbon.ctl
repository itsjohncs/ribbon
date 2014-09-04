VERSION 5.00
Begin VB.UserControl BHRibbon 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   PropertyPages   =   "BHRibbon.ctx":0000
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   130
   Begin VB.Timer tmrMouseover 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   360
      Top             =   720
   End
   Begin VB.PictureBox picCat 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   0
      Left            =   0
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
      Begin VB.PictureBox picBut 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   990
         Index           =   0
         Left            =   120
         ScaleHeight     =   66
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   3
         Top             =   60
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox picDlg 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   1560
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "BHRibbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'|||||||||||||||||||||||||||||||||||||
'|||          BHRibbon V1          |||
'|||-------------------------------|||
'|||    BHRibbon was created by    |||
'|||          -Brownhead-          |||
'|||   (brownhead@brownhead.com)   |||
'|||      (www.brownhead.com)      |||
'|||||||||||||||||||||||||||||||||||||

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long

Event TabClick(Index As Integer, Caption As String)
Event ButClick(Index As Integer, Caption As String, Category As Integer)
Event DlgClick(Index As Integer)

Private Type POINTAPI 'Used to retrieve the cursor coordinates
    x As Long
    y As Long
End Type
Private Type RECT 'Used with the DrawText API
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type UserButton
    Index As Integer 'Postion in tab (Relative to the tab. If theres 4 buttons in the tab and this is the fourth one, its number 4)
    ID As Integer 'Position in category (Relative to the category. If its the first in the category, than it equals 1, no matter how many buttons there are)
    Category As Integer 'The Category ID its in (Relative to the current tab. If theres 2 categories in the tab and its in the second category, its 2 no matter how many categories there are in all)
    AbsoluteCat As Integer 'The category ID according to the tCat's indext
    Caption As String
    Icon As Integer
    More As Boolean 'If this is true it will make a small arrow underneath the caption
    Tooltip As String
    Width As Integer
    Tab As Integer
End Type
Private Type Category
    AbsoluteID As Integer 'The actual number of the category. Reserved for tCurCat array
    ID As Integer 'Position in tab (Relative to the tab)
    Caption As String
    Tab As Integer
    Dialog As Boolean 'Will supply an added button in the caption
    Width As Integer
End Type
Private Type Style 'Why use a type in this case? Because its ALOT easier to handle
    TabL As IPictureDisp 'Left part of the tabs
    TabL_ON As IPictureDisp 'The left part of the tab during a mouseover
    TabM As IPictureDisp 'Middle part of the tabs
    TabM_ON As IPictureDisp
    TabR As IPictureDisp 'Right part of the tabs
    TabR_ON As IPictureDisp
    ButL As IPictureDisp 'Left part of the button
    ButL_ON As IPictureDisp
    ButL_DN As IPictureDisp 'The left part of the button during a mouseclick
    ButM As IPictureDisp 'Middle part of the button
    ButM_ON As IPictureDisp
    ButM_DN As IPictureDisp 'The middle part of the button during a mouseclick
    ButR As IPictureDisp 'Right part of the button
    ButR_ON As IPictureDisp
    ButR_DN As IPictureDisp 'The right part of the button during mouseclick
    CatL As IPictureDisp 'The left part of the category
    CatL_ON As IPictureDisp
    CatM As IPictureDisp 'The middle part of the category
    CatM_ON As IPictureDisp
    CatR As IPictureDisp 'The right part of the cateogry
    CatR_ON As IPictureDisp
    Dlg As IPictureDisp 'The dialog button's picture
    Dlg_CON As IPictureDisp 'The dialog button when the parent category is being moused over
    Dlg_ON As IPictureDisp 'The dialog button when the actaul button is being moused over
    BgL As IPictureDisp 'The left part of the background
    BgM As IPictureDisp 'The middle part of the background
    FormBG As IPictureDisp 'The background of the parent form
    More As IPictureDisp
    More_ON As IPictureDisp
    TabBg As OLE_COLOR 'The color of the Tabs area (The strip of color that the tabs run throuhg)
    TabColor As OLE_COLOR 'The tab text color
    TabColor_SEL As OLE_COLOR 'The tab text color when selected
    CatColor As OLE_COLOR 'The catagory text color
    ButColor As OLE_COLOR 'The button text color
End Type
Public Enum Styles 'The default styles. Custom ones are made via the property page
    Black = 1
    Blue = 2
    Silver = 3
End Enum

Dim tBut() As UserButton, tCat() As Category, sTab() As String, tCurCat() As Category, tCurBut() As UserButton, lOver As Long, iCurrentTab As Integer, tStyle As Style
Dim rCharSet As Integer, rRightLeft As Boolean, rFontFamily As String, rStyle As Integer, rBut As String, rCat As String, rTab As String, rImageList As ImageList 'Holds the values of all the properties

'//Friend Procedures :: These are only accessible in the property pages\\
Friend Property Get pIcon(Index As Integer) As IPictureDisp
Set pIcon = rImageList.ListImages(Index).Picture
End Property
Friend Property Get Tabs() As String
Dim iCount As Integer
If (Not sTab) = True Then Exit Sub
For iCount = 0 To UBound(sTab)
    Tabs = Tabs & sTab(iCount) & Chr(1)
Next iCount
Tabs = Left(Tabs, Len(Tabs) - 1)
End Property
Friend Property Let Tabs(newValue As String)
Dim aHold() As String, iCount As Integer
Erase sTab
rTab = newValue
aHold = Split(newValue, Chr(1))
For iCount = 0 To UBound(aHold)
    If (Len(aHold(iCount)) > 0) Then AddTab aHold(iCount)
Next iCount
PropertyChanged "Tabs"
End Property
Friend Property Get Cats() As String
Dim iCount As Integer
For iCount = 0 To UBound(tCat)
    With tCat(iCount)
        Cats = Cats & .Tab & Chr(1) & .Caption & Chr(1) & IIf(.Dialog, "T", "F") & Chr(2)
    End With
Next iCount
Cats = Left(Cats, Len(Cats) - 1)
End Property
Friend Property Let Cats(newValue As String)
Dim aHold() As String, aHold2() As String, iCount As Integer
Erase tCat
rCat = newValue
aHold = Split(newValue, Chr(1))
For iCount = 0 To UBound(aHold)
    If (Len(aHold(iCount)) > 0) Then
        aHold2 = Split(aHold(iCount), Chr(2))
        AddCat aHold2(0), aHold2(1), aHold2(2) = "T"
    End If
Next iCount
PropertyChanged "Cats"
End Property
Friend Property Get Buts() As String
Dim iCount As Integer
For iCount = 0 To UBound(tBut)
    With tBut(iCount)
        Buts = Buts & .Tab & Chr(1) & .Category & Chr(1) & .Icon & Chr(1) & .Caption & Chr(1) & .Tooltip & Chr(1) & .More & Chr(2)
    End With
Next iCount
Buts = Left(Buts, Len(Buts) - 1)
End Property
Friend Property Let Buts(newValue As String)
Dim aHold() As String, aHold2() As String, iCount As Integer
Erase tBut
rBut = Buts
aHold = Split(newValue, Chr(1))
For iCount = 0 To UBound(sBut)
    If (Len(aHold(iCount)) > 0) Then
        aHold2 = Split(aHold(iCount), Chr(2))
        AddBut aHold2(0), aHold2(1), aHold2(2), aHold2(3), aHold2(4), aHold2(5) = "T"
    End If
Next iCount
PropertyChanged "Buts"
End Property
'\\Friend Procedures :: These are only accessible in the property pages//

'//Properties\\
Public Property Let ImageList(newValue As ImageList)
Set rImageList = newValue
End Property
Public Property Get SelectedTab() As Integer
SelectedTab = iCurrentTab
End Property
Public Property Let RightLeft(ByVal newValue As Boolean)
rRightLeft = newValue
PropertyChanged "RightLeft"
Refresh
End Property
Public Property Get RightLeft() As Boolean
RightLeft = rRightLeft
End Property
Public Property Let CharSet(ByVal newValue As Integer)
Dim tControl As Control
rCharSet = newValue
For Each tControl In UserControl.Controls
    If (Left(tControl.Name, 3) = "pic") Then
        tControl.Font.CharSet = newValue
        PropertyChanged "CharSet"
    End If
Next tControl
Refresh
End Property
Public Property Get CharSet() As Integer
CharSet = rCharSet
End Property
Public Property Let FontFamily(ByVal newValue As String)
Dim tControl As Control
If FontExists(newValue) Then
    rFontFamily = newValue
    For Each tControl In UserControl.Controls
        If (Left(tControl.Name, 3) = "pic") Then
            tControl.Font.Name = newValue
            PropertyChanged "FontFamily"
        End If
    Next tControl
End If
Refresh
End Property
Public Property Get FontFamily() As String
FontFamily = rFontFamily
End Property
Public Property Let Style(ByVal newValue As Styles)
rStyle = newValue
With tStyle 'This big block will retrieve the images and such needed for the styles. If you need to know what each property means, check the type declarations, I commented it in excess to make it easy ofr you to figure everything out
    If (rStyle <> 3) Then 'The silver theme does not have a background
        Set .FormBG = LoadResPicture(1 + rStyle * 100, vbResBitmap)
    End If
    
    Set .ButR_ON = LoadResPicture(2 + rStyle * 100, vbResBitmap)
    Set .ButM_ON = LoadResPicture(3 + rStyle * 100, vbResBitmap)
    Set .ButL_ON = LoadResPicture(4 + rStyle * 100, vbResBitmap)
    Set .ButR_DN = LoadResPicture(5 + rStyle * 100, vbResBitmap)
    Set .ButM_DN = LoadResPicture(6 + rStyle * 100, vbResBitmap)
    Set .ButL_DN = LoadResPicture(7 + rStyle * 100, vbResBitmap)
    
    Set .More = LoadResPicture(8 + rStyle * 100, vbResBitmap)
    Set .More_ON = LoadResPicture(9 + rStyle * 100, vbResBitmap)
    
    Set .TabR_ON = LoadResPicture(10 + rStyle * 100, vbResBitmap)
    Set .TabM_ON = LoadResPicture(11 + rStyle * 100, vbResBitmap)
    Set .TabL_ON = LoadResPicture(12 + rStyle * 100, vbResBitmap)
    Set .TabR = LoadResPicture(13 + rStyle * 100, vbResBitmap)
    Set .TabM = LoadResPicture(14 + rStyle * 100, vbResBitmap)
    Set .TabL = LoadResPicture(15 + rStyle * 100, vbResBitmap)
    
    Set .CatR_ON = LoadResPicture(16 + rStyle * 100, vbResBitmap)
    Set .CatM_ON = LoadResPicture(17 + rStyle * 100, vbResBitmap)
    Set .CatL_ON = LoadResPicture(18 + rStyle * 100, vbResBitmap)
    Set .CatR = LoadResPicture(19 + rStyle * 100, vbResBitmap)
    Set .CatM = LoadResPicture(20 + rStyle * 100, vbResBitmap)
    Set .CatL = LoadResPicture(21 + rStyle * 100, vbResBitmap)
    
    Set .Dlg_ON = LoadResPicture(22 + rStyle * 100, vbResBitmap)
    Set .Dlg_CON = LoadResPicture(23 + rStyle * 100, vbResBitmap)
    Set .Dlg = LoadResPicture(24 + rStyle * 100, vbResBitmap)
   
    Set .BgL = LoadResPicture(26 + rStyle * 100, vbResBitmap)
    Set .BgM = LoadResPicture(27 + rStyle * 100, vbResBitmap)

    .CatColor = Val(LoadResString(1 + rStyle * 100))
    .ButColor = Val(LoadResString(2 + rStyle * 100))
    .TabColor = Val(LoadResString(3 + rStyle * 100))
    .TabColor_SEL = Val(LoadResString(4 + rStyle * 100))
    .TabBg = Val(LoadResString(5 + rStyle * 100))
End With
PropertyChanged "Style"
Refresh
End Property
Public Property Get Style() As Styles
Style = rStyle
End Property
'\\Properties//

'//Usercontrol Events\\
'//Usercontrol Property Events\\
Private Sub UserControl_InitProperties()
Style = Black
rRightLeft = False
FontFamily = "Verdana"
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
With PropBag
    Style = .ReadProperty("Style", 1)
    rRightLeft = .ReadProperty("RightLeft", False)
    FontFamily = .ReadProperty("FontFamily", "Verdana")
    rCharSet = .ReadProperty("CharSet")
End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Style", rStyle, 1
    .WriteProperty "RightLeft", rRightLeft, False
    .WriteProperty "FontFamily", rFontFamily, "Verdana"
    .WriteProperty "Tabs", rTab
    .WriteProperty "Cats", rCat
    .WriteProperty "Buts", rBut
    .WriteProperty "CharSet", rCharSet
End With
End Sub
'\\Usercontrol Property Events//
Private Sub UserControl_Show()
If (tStyle.ButColor > 0) Then Refresh
End Sub
Private Sub UserControl_Resize()
UserControl.Height = 1740
If (tStyle.ButColor <> 0) Then Refresh 'Refresh the control
End Sub
'\\Usercontrol Events//

'//picBut Events\\
Private Sub picBut_Click(Index As Integer)
RaiseEvent ButClick(Index, tCurBut(Index).Caption, tCurBut(Index).Category)
End Sub
Private Sub picBut_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim rHold As RECT
With picBut(Index)
    .PaintPicture tStyle.ButL_DN, 0, 0
    .PaintPicture tStyle.ButM_DN, 3, 0, .Width - 6
    .PaintPicture tStyle.ButR_DN, .Width - 3, 0
    BltPic rImageList.ListImages(tCurBut(Index).Icon).Picture, .hdc, .Width / 2 - ScaleX(rImageList.ListImages(tCurBut(Index).Icon).Picture.Width, vbHimetric, vbPixels) / 2, 4, IIf(rImageList.ListImages(tCurBut(Index).Icon).Tag <> "", Val(rImageList.ListImages(tCurBut(Index).Icon).Tag), vbNull) 'Draws the icon\
    SetRect rHold, 2, 0, .Width - 2, 18 + ScaleY(rImageList.ListImages(tCurBut(Index).Icon).Picture.Height, vbHimetric, vbPixels) 'A quick way to assign values to a RECT structure
    DrawText picBut(Index).hdc, tCurBut(Index).Caption, Len(tCurBut(Index).Caption), rHold, IIf(rRightLeft, &H20000 Or &H2, &H1) Or &H8 Or &H20 'Draws the text, taking into account whether RightLeft is enabled or not
End With
End Sub
Private Sub picBut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim rHold As RECT, iCount As Integer
If (picBut(Index).hWnd = lOver) Then Exit Sub
For iCount = 0 To picBut.UBound
    With picBut(iCount)
        If .Visible Then
            If (tCurBut(iCount).Category = tCurBut(Index).Category) Then
                .PaintPicture tStyle.CatM_ON, 0, -4, .Width 'Paint the buttons background, which is the category background offset slightly
                BltPic rImageList.ListImages(tCurBut(iCount).Icon).Picture, .hdc, .Width / 2 - ScaleX(rImageList.ListImages(tCurBut(iCount).Icon).Picture.Width, vbHimetric, vbPixels) / 2, 4, IIf(rImageList.ListImages(tCurBut(iCount).Icon).Tag <> "", Val(rImageList.ListImages(tCurBut(iCount).Icon).Tag), vbNull) 'Draws the icon\
                SetRect rHold, 2, 0, .Width - 2, 18 + ScaleY(rImageList.ListImages(tCurBut(iCount).Icon).Picture.Height, vbHimetric, vbPixels) 'A quick way to assign values to a RECT structure
                DrawText picBut(iCount).hdc, tCurBut(iCount).Caption, Len(tCurBut(iCount).Caption), rHold, IIf(rRightLeft, &H20000 Or &H2, &H1) Or &H8 Or &H20 'Draws the text, taking into account whether RightLeft is enabled or not
                If tCurBut(iCount).More Then BltPic tStyle.More_ON, .hdc, .Width / 2 - 2, rHold.Bottom + 3, &HFFFFFF
            End If
        End If
    End With
Next iCount
With picBut(Index)
    .PaintPicture tStyle.ButL_ON, 0, 0
    .PaintPicture tStyle.ButM_ON, 3, 0, .Width - 6
    .PaintPicture tStyle.ButR_ON, .Width - 3, 0
    BltPic rImageList.ListImages(tCurBut(Index).Icon).Picture, .hdc, .Width / 2 - ScaleX(rImageList.ListImages(tCurBut(Index).Icon).Picture.Width, vbHimetric, vbPixels) / 2, 4, IIf(rImageList.ListImages(tCurBut(Index).Icon).Tag <> "", Val(rImageList.ListImages(tCurBut(Index).Icon).Tag), vbNull) 'Draws the icon\
    SetRect rHold, 2, 0, .Width - 2, 18 + ScaleY(rImageList.ListImages(tCurBut(Index).Icon).Picture.Height, vbHimetric, vbPixels) 'A quick way to assign values to a RECT structure
    DrawText picBut(Index).hdc, tCurBut(Index).Caption, Len(tCurBut(Index).Caption), rHold, IIf(rRightLeft, &H20000 Or &H2, &H1) Or &H8 Or &H20 'Draws the text, taking into account whether RightLeft is enabled or not
    If tCurBut(Index).More Then BltPic tStyle.More_ON, .hdc, .Width / 2 - 2, rHold.Bottom + 3, &HFFFFFF
    If (lOver = 0) Then lOver = .hWnd
End With
With picCat(tCurBut(Index).Category)
    .PaintPicture tStyle.CatL_ON, 0, 0
    .PaintPicture tStyle.CatM_ON, 4, 0, .Width - 8
    .PaintPicture tStyle.CatR_ON, .Width - 4, 0
    SetRect rHold, 2, 0, .Width - 3 - IIf(tCurCat(tCurBut(Index).Category).Dialog, 15, 0), .Height - 4 'A quick way to assign values to a RECT structure
    DrawText .hdc, tCurCat(tCurBut(Index).Category).Caption, Len(tCurCat(tCurBut(Index).Category).Caption), rHold, IIf(rRightLeft, &H20000 Or &H2, &H1) Or &H8 Or &H20 'Draws the text, taking into account whether RightLeft is enabled or not
End With
picDlg(tCurBut(Index).Category).PaintPicture tStyle.Dlg_CON, 0, 0
tmrMouseover.Enabled = True
End Sub
Private Sub picBut_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim rHold As RECT
With picBut(Index)
    .PaintPicture tStyle.ButL_ON, 0, 0
    .PaintPicture tStyle.ButM_ON, 3, 0, .Width - 6
    .PaintPicture tStyle.ButR_ON, .Width - 3, 0
    BltPic rImageList.ListImages(tCurBut(Index).Icon).Picture, .hdc, .Width / 2 - ScaleX(rImageList.ListImages(tCurBut(Index).Icon).Picture.Width, vbHimetric, vbPixels) / 2, 4, IIf(rImageList.ListImages(tCurBut(Index).Icon).Tag <> "", Val(rImageList.ListImages(tCurBut(Index).Icon).Tag), vbNull) 'Draws the icon\
    SetRect rHold, 2, 0, .Width - 2, 18 + ScaleY(rImageList.ListImages(tCurBut(Index).Icon).Picture.Height, vbHimetric, vbPixels) 'A quick way to assign values to a RECT structure
    DrawText picBut(Index).hdc, tCurBut(Index).Caption, Len(tCurBut(Index).Caption), rHold, IIf(rRightLeft, &H20000 Or &H2, &H1) Or &H8 Or &H20 'Draws the text, taking into account whether RightLeft is enabled or not
    If tCurBut(Index).More Then BltPic tStyle.More_ON, .hdc, .Width / 2 - 2, rHold.Bottom + 3, &HFFFFFF
End With
End Sub
'\\picBut Events//

'//picCat Events\\
Private Sub picCat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim rHold As RECT
If (picCat(Index).hWnd = lOver) Then Exit Sub
With picCat(Index)
    .PaintPicture tStyle.CatL_ON, 0, 0
    .PaintPicture tStyle.CatM_ON, 4, 0, .Width - 8
    .PaintPicture tStyle.CatR_ON, .Width - 4, 0
    SetRect rHold, 2, 0, .Width - 3 - IIf(tCurCat(Index).Dialog, 15, 0), .Height - 4 'A quick way to assign values to a RECT structure
    DrawText .hdc, tCurCat(Index).Caption, Len(tCurCat(Index).Caption), rHold, IIf(rRightLeft, &H20000 Or &H2, &H1) Or &H8 Or &H20 'Draws the text, taking into account whether RightLeft is enabled or not
    If (lOver = 0) Then lOver = .hWnd
End With
For iCount = 0 To picBut.UBound
    With picBut(iCount)
        If .Visible Then 'If the button is visible
            If (tCurBut(iCount).Category = Index) Then
                .PaintPicture tStyle.CatM_ON, 0, -4, .Width 'Paint the buttons background, which is the category background offset slightly
                BltPic rImageList.ListImages(tCurBut(iCount).Icon).Picture, .hdc, .Width / 2 - ScaleX(rImageList.ListImages(tCurBut(iCount).Icon).Picture.Width, vbHimetric, vbPixels) / 2, 4, IIf(rImageList.ListImages(tCurBut(iCount).Icon).Tag <> "", Val(rImageList.ListImages(tCurBut(iCount).Icon).Tag), vbNull) 'Draws the icon\
                SetRect rHold, 2, 0, .Width - 2, 18 + ScaleY(rImageList.ListImages(tCurBut(iCount).Icon).Picture.Height, vbHimetric, vbPixels) 'A quick way to assign values to a RECT structure
                DrawText picBut(iCount).hdc, tCurBut(iCount).Caption, Len(tCurBut(iCount).Caption), rHold, IIf(rRightLeft, &H20000 Or &H2, &H1) Or &H8 Or &H20 'Draws the text, taking into account whether RightLeft is enabled or not
                If tCurBut(iCount).More Then BltPic tStyle.More_ON, .hdc, .Width / 2 - 2, rHold.Bottom + 3, &HFFFFFF
            End If
        End If
    End With
Next iCount
picDlg(Index).PaintPicture tStyle.Dlg_CON, 0, 0
tmrMouseover.Enabled = True
End Sub
'\\picCat Events//

'//picDlg Events\\
Private Sub picDlg_Click(Index As Integer)
RaiseEvent DlgClick(Index)
End Sub
Private Sub picDlg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim rHold As RECT
If (picBut(Index).hWnd = lOver) Then Exit Sub
If (lOver = 0) Then lOver = picDlg(Index).hWnd
picDlg(Index).PaintPicture tStyle.Dlg_ON, 0, 0
With picCat(Index)
    .PaintPicture tStyle.CatL_ON, 0, 0
    .PaintPicture tStyle.CatM_ON, 4, 0, .Width - 8
    .PaintPicture tStyle.CatR_ON, .Width - 4, 0
    SetRect rHold, 2, 0, .Width - 3 - IIf(tCurCat(Index).Dialog, 15, 0), .Height - 4 'A quick way to assign values to a RECT structure
    DrawText .hdc, tCurCat(Index).Caption, Len(tCurCat(Index).Caption), rHold, IIf(rRightLeft, &H20000 Or &H2, &H1) Or &H8 Or &H20 'Draws the text, taking into account whether RightLeft is enabled or not
    If (lOver = 0) Then lOver = picDlg(Index).hWnd
End With
For iCount = 0 To picBut.UBound
    With picBut(iCount)
        If .Visible Then 'If the button is visible
            If (tCurBut(iCount).Category = Index) Then
                .PaintPicture tStyle.CatM_ON, 0, -4, .Width 'Paint the buttons background, which is the category background offset slightly
                BltPic rImageList.ListImages(tCurBut(iCount).Icon).Picture, .hdc, .Width / 2 - ScaleX(rImageList.ListImages(tCurBut(iCount).Icon).Picture.Width, vbHimetric, vbPixels) / 2, 4, IIf(rImageList.ListImages(tCurBut(iCount).Icon).Tag <> "", Val(rImageList.ListImages(tCurBut(iCount).Icon).Tag), vbNull) 'Draws the icon\
                SetRect rHold, 2, 0, .Width - 2, 18 + ScaleY(rImageList.ListImages(tCurBut(iCount).Icon).Picture.Height, vbHimetric, vbPixels) 'A quick way to assign values to a RECT structure
                DrawText picBut(iCount).hdc, tCurBut(iCount).Caption, Len(tCurBut(iCount).Caption), rHold, IIf(rRightLeft, &H20000 Or &H2, &H1) Or &H8 Or &H20 'Draws the text, taking into account whether RightLeft is enabled or not
                If tCurBut(iCount).More Then BltPic tStyle.More_ON, .hdc, .Width / 2 - 2, rHold.Bottom + 3, &HFFFFFF
            End If
        End If
    End With
Next iCount
tmrMouseover.Enabled = True
End Sub
'\\picDlg Events//

'//picTab Events\\
Private Sub picTab_Click(Index As Integer)
iCurrentTab = Index
Refresh
RaiseEvent TabClick(Index, sTab(Index))
End Sub
Private Sub picTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim rHold As RECT
If (Index = iCurrentTab Or picTab(Index).hWnd = lOver) Then Exit Sub
With picTab(Index)
    .PaintPicture tStyle.TabL_ON, 0, 0
    .PaintPicture tStyle.TabM_ON, 10, 0, .Width - 20
    .PaintPicture tStyle.TabR_ON, .Width - 10, 0
    SetRect rHold, 5, 0, .Width - 5, 24 'A quick way to assign values to a RECT structure
    DrawText picTab(Index).hdc, sTab(Index), Len(sTab(Index)), rHold, IIf(rRightLeft, &H20000 Or &H2, &H1) Or &H4 Or &H20 'Draws the text, taking into account whether RightLeft is enabled or not
    If (lOver = 0) Then lOver = .hWnd
End With
tmrMouseover.Enabled = True
End Sub
'\\picTab Events//

'//tmrMouserover Events\\
Private Sub tmrMouseover_Timer()
Dim tHold As POINTAPI
GetCursorPos tHold
If (WindowFromPoint(tHold.x, tHold.y) <> lOver) Then
    Refresh
    lOver = 0
    tmrMouseover.Enabled = False
End If
End Sub
'\\tmrMouseover Events//

'//Public Functions\\
Public Sub Clear() 'Erases all the Buttons, Categories and Tabs
Erase sTab
Erase tBut
Erase tCat
End Sub
Public Sub AddTab(ByVal Caption As String)
If (Not sTab) = True Then ReDim sTab(0) Else ReDim Preserve sTab(UBound(sTab) + 1)
sTab(UBound(sTab)) = Caption
End Sub
Public Sub AddCat(ByVal iTab As Integer, ByVal Caption As String, Optional ByVal Dialog As Boolean)
Dim iCount As Integer, iHold As Integer
If (Not tCat) = True Then ReDim tCat(0) Else ReDim Preserve tCat(UBound(tCat) + 1)
With tCat(UBound(tCat))
    For iCount = 0 To UBound(tCat)
        If (tCat(iCount).Tab = iTab) Then iHold = iHold + 1
        If (iHold > 0) And (iCount = UBound(tCat) Or tCat(iCount).Tab <> iTab) Then
            .AbsoluteID = iCount
            .ID = iHold
        End If
    Next iCount
    .Caption = Caption
    .Dialog = Dialog
    .Tab = iTab
End With
End Sub
Public Sub AddBut(ByVal iTab As Integer, ByVal Category As Integer, ByVal Icon As Integer, ByVal Caption As String, Optional ByVal Tooltip As String, Optional ByVal More As Boolean)
Dim iCount As Integer, iHold As Integer, iHold2 As Integer
If (Not tBut) = True Then ReDim tBut(0) Else ReDim Preserve tBut(UBound(tBut) + 1)
With tBut(UBound(tBut))
    For iCount = 0 To UBound(tCat)
        If (iHold = Category) Then .AbsoluteCat = iCount
        If (tCat(iCount).Tab = iTab) Then iHold = iHold + 1
    Next iCount
    .Caption = Caption
    .Category = Category
    If (Icon < 1 Or Icon > rImageList.ListImages.Count) Then .Icon = 1 Else .Icon = Icon
    .More = More
    .Tab = iTab
    .Tooltip = Tooltip
    iHold = 0
    For iCount = 0 To UBound(tBut)
        If (tBut(iCount).Category = Category And tBut(iCount).Tab = iTab) Then iHold = iHold + 1
        If (tBut(iCount).Tab = iTab) Then iHold2 = iHold2 + 1
        If (tBut(iCount).Category = Category And tBut(iCount).Tab = iTab) And (iHold > 0 Or iCount = UBound(tBut)) Then
            .ID = iHold - 1
            .Index = iHold2 - 1
        End If
    Next iCount
End With
End Sub
Public Sub Refresh() 'Refreshes the controls and draws everything. Takes 10 to 40 MS to complete. If you need to make it faster swap out all the PaintPicture functions with BitBlt's and StrethBlt's
On Error Resume Next
Dim iCount As Integer, iCount2 As Integer, rHold As RECT, iHold As Integer, iHold2() As Integer, tPic As PictureBox
UserControl.BackColor = tStyle.TabBg 'Sets the backcolor of the control to the styles background color
PaintPicture tStyle.BgL, 0, -26 'Sets the left part of the background
PaintPicture tStyle.BgM, 7, -26, UserControl.Width / Screen.TwipsPerPixelX 'Sets the middle part of the background
'//Tabs\\
If (Not sTab) = True Then Exit Sub 'Checks first to see if the array is initialized or not
For iCount = 0 To IIf(UBound(sTab) > picTab.UBound, UBound(sTab), picTab.UBound) 'Loops through all the tabs and detirmines whether there needs to be more or less
    If (iCount > picTab.UBound And iCount <= UBound(sTab)) Then 'If there is no picturebox made for this
        Load picTab(iCount) 'Loads another tab
        picTab(iCount).Visible = True 'I never did figure out why Microsoft made the newly loaded tabs invisible.. w/e.. just 1 extra line of code
    ElseIf (iCount <= picTab.UBound And iCount > UBound(sTab)) Then 'If the picture box is made but shouldn't be
        If (iCount = 0) Then picTab(iCount).Visible = False Else Unload picTab(iCount) 'Unloads the tab (Or hides it if unloading is impossible... you cannot unload the first item of an object array)
    Else
        picTab(iCount).Visible = True 'I ahve the initial pictureboxes visibility set to false at runtime, this makes them visible again if need be
    End If
Next iCount
iHold = 7 'Sets the initial offset. So the first tab isin the right place
For iCount = 0 To picTab.UBound 'Loops through all of the tabs and draws there content
    With picTab(iCount)
        .Width = .TextWidth(sTab(iCount)) + 24 'Sets the size of the tab
        .Left = Abs(iHold - IIf(rRightLeft, UserControl.ScaleWidth, 0)) - IIf(rRightLeft, .Width, 0)
        iHold = iHold + .Width + IIf(iHold = 7, 0, 2)
        If (iCount = iCurrentTab) Then 'If its the current tab...
            .ForeColor = tStyle.TabColor_SEL 'Sets the text color
            .PaintPicture tStyle.TabL, 0, 0 'Paint the left part
            .PaintPicture tStyle.TabM, 10, 0, .Width - 20 'Paint the middle part
            .PaintPicture tStyle.TabR, .Width - 10, 0 'Paint the right part
        Else
            .ForeColor = tStyle.TabColor 'Sets the text color
            .BackColor = tStyle.TabBg 'Sets the background, there is no image bg for this
        End If
        SetRect rHold, 5, 0, .Width - 5, 24 'A quick way to assign values to a RECT structure
        DrawText picTab(iCount).hdc, sTab(iCount), Len(sTab(iCount)), rHold, IIf(rRightLeft, &H20000 Or &H2, &H1) Or &H4 Or &H20 'Draws the text, taking into account whether RightLeft is enabled or not
    End With
Next iCount
'\\Tabs//
'//Buttons & Categories :: Only sets the widths (Necassery for next block)\\
If Not (Not tBut) = True Then 'If the array is initialized. Do not try to streamline this line, I know I have two not statements.. everything's needed
    For iCount = 0 To UBound(tBut) 'Loops through every button
        If (tBut(iCount).Tab = iCurrentTab) Then
            If (tBut(iCount).ID = 0) Then tCat(tBut(iCount).AbsoluteCat).Width = 0
            tBut(iCount).Width = IIf(picBut(0).TextWidth(tBut(iCount).Caption) + 10 > 24, picBut(0).TextWidth(tBut(iCount).Caption) + 10, 24) 'Sets the button's width
            If (tCat(tBut(iCount).AbsoluteCat).Width = 0) Then tCat(tBut(iCount).AbsoluteCat).Width = tCat(tBut(iCount).AbsoluteCat).Width + 6
            tCat(tBut(iCount).AbsoluteCat).Width = tCat(tBut(iCount).AbsoluteCat).Width + tBut(iCount).Width + 4 'Adds the button to the width of the category
        End If
     Next iCount
End If
'\\Buttons & Categories :: Only sets the widths (Necassery for next block)//
'//Categories & Dialog Boxes\\
If (Not tCat) = True Then Exit Sub 'If the category array hasn't been initialized than exit the sub
Erase tCurCat 'Erases the tCurCat array
For iCount = 0 To UBound(tCat) 'Places the current cateogories into a variable
    If (tCat(iCount).Tab = iCurrentTab) Then 'If it should be displayed
        If (Not tCurCat) = True Then ReDim tCurCat(0) As Category Else ReDim Preserve tCurCat(UBound(tCurCat) + 1) As Category  'Add an element to the array
        tCurCat(UBound(tCurCat)) = tCat(iCount) 'Sets the element
    End If
Next iCount
If (Not tCurCat) = True Then 'If there are no categories to be displayed, hide everything and exit the sub
    For iCount = 0 To picCat.UBound
        picCat(iCount).Visible = False
    Next iCount
    Exit Sub
End If
For iCount = 0 To IIf(UBound(tCurCat) > picCat.UBound, UBound(tCurCat), picCat.UBound)
    If (iCount > UBound(tCurCat)) Then
        picCat(iCount).Visible = False 'It error if you try to unload it,
        picDlg(iCount).Visible = False 'I'm actually not sure why.
    Else
        If (iCount > picCat.UBound) Then 'If a picturebox with this index has not been created yet
            Load picCat(iCount) 'Create the Category
            Load picDlg(iCount) 'Create the corresponding dialog button
            SetParent picDlg(iCount).hWnd, picCat(iCount).hWnd 'Moves the dialog button into the correct category
        End If
        picCat(iCount).Visible = True 'Make the category visible
        picDlg(iCount).Visible = tCurCat(iCount).Dialog 'Make the dialog button visible or invisible depending on the cateogry
    End If
Next iCount
iHold = 7 'Sets the intial offset. So that the first cateogy is 8 pixels away form the side of the control
For iCount = 0 To UBound(tCurCat)
    With picCat(iCount)
        .Width = IIf(.TextWidth(tCurCat(iCount).Caption) + 24 > tCurCat(iCount).Width, .TextWidth(tCat(iCount).Caption) + 24, tCurCat(iCount).Width)
        .Left = Abs(iHold - IIf(rRightLeft, UserControl.ScaleWidth, 0)) - IIf(rRightLeft, .Width, 0) 'Aligns it correctly
        iHold = iHold + .Width + IIf(iHold = 7, 0, 1)  'This marks the place so it doesn't overlap other categories
        .ForeColor = tStyle.CatColor 'Sets the text color
        .PaintPicture tStyle.CatL, 0, 0 'Paints the left part of the category
        .PaintPicture tStyle.CatM, 4, 0, .Width - 8 'Paints the middle part
        .PaintPicture tStyle.CatR, .Width - 4, 0 'Paints the right part
        SetRect rHold, 2, 0, .Width - 3 - IIf(tCurCat(iCount).Dialog, 15, 0), .Height - 4 'A quick way to assign values to a RECT structure
        DrawText picCat(iCount).hdc, tCurCat(iCount).Caption, Len(tCurCat(iCount).Caption), rHold, IIf(rRightLeft, &H20000 Or &H2, &H1) Or &H8 Or &H20 'Draws the text, taking into account whether RightLeft is enabled or not
    End With
    With picDlg(iCount)
        .Move picCat(iCount).Width - 3 - .Width, picCat(iCount).Height - 2 - .Width
        .PaintPicture tStyle.Dlg, 0, 0
    End With
Next iCount
'\\Categories & Dialog Boxes//
'//Buttons\\
If (Not tBut) = True Then Exit Sub 'If the tBut array hasn't been initialized than exit the method
Erase tCurBut 'Erase the tCurBut array
For iCount = 0 To UBound(tBut) 'Places the current buttons into a variable
    If (tBut(iCount).Tab = iCurrentTab) Then 'If it should be displayed
        If (Not tCurBut) = True Then ReDim tCurBut(0) As UserButton Else ReDim Preserve tCurBut(UBound(tCurBut) + 1) As UserButton  'Add an element to the array
        tCurBut(UBound(tCurBut)) = tBut(iCount) 'Sets the element
    End If
Next iCount
If (Not tCurBut) = True Then 'If there are no current buttons to display hide everything and exit the sub
    For iCount = 0 To picBut.UBound
        picBut(iCount).Visible = False
    Next iCount
    Exit Sub
End If
For iCount = 0 To IIf(UBound(tCurBut) > picBut.UBound, UBound(tCurBut), picBut.UBound)
    If (iCount > UBound(tCurBut)) Then
        picBut(iCount).Visible = False 'It error if you try to unload it,
    Else
        If (iCount > picBut.UBound) Then 'If a picturebox with this index has not been created yet
            Load picBut(iCount) 'Create the button
        End If
        SetParent picBut(iCount).hWnd, picCat(tCurBut(iCount).Category).hWnd 'Moves the button into the correct category
        picBut(iCount).Visible = True 'Make the button visible
    End If
Next iCount
ReDim iHold2(UBound(tCat)) As Integer 'I use this to mark my place in each category
For iCount = 0 To UBound(tCurBut)
    With picBut(iCount)
        .ToolTipText = tCurBut(iCount).Tooltip
        .Width = IIf(tCurBut(iCount).Width > 24, tCurBut(iCount).Width, 24) 'Sets the size of the Button
        If rRightLeft Then .Left = picCat(tCurBut(iCount).Category).Width - .Width - iHold2(tCurBut(iCount).Category) - IIf(tCurBut(iCount).ID = 0, 5, 0) Else .Left = iHold2(tCurBut(iCount).Category) + IIf(tCurBut(iCount).ID = 0, 4, 0)
        iHold2(tCurBut(iCount).Category) = iHold2(tCurBut(iCount).Category) + .Width + 2 + IIf(tCurBut(iCount).ID = 0, 8, 0) 'Marks my place for the current category
        .PaintPicture tStyle.CatM, 0, -4, .Width 'Paint the buttons background, which is the category background offset slightly
        BltPic rImageList.ListImages(tCurBut(iCount).Icon).Picture, .hdc, .Width / 2 - ScaleX(rImageList.ListImages(tCurBut(iCount).Icon).Picture.Width, vbHimetric, vbPixels) / 2, 4, IIf(rImageList.ListImages(tCurBut(iCount).Icon).Tag <> "", Val(rImageList.ListImages(tCurBut(iCount).Icon).Tag), vbNull) 'Draws the icon
        .ForeColor = tStyle.ButColor
        SetRect rHold, 2, 0, .Width - 2, 18 + ScaleY(rImageList.ListImages(tCurBut(iCount).Icon).Picture.Height, vbHimetric, vbPixels) 'A quick way to assign values to a RECT structure
        DrawText picBut(iCount).hdc, tCurBut(iCount).Caption, Len(tCurBut(iCount).Caption), rHold, IIf(rRightLeft, &H20000 Or &H2, &H1) Or &H8 Or &H20 'Draws the text, taking into account whether RightLeft is enabled or not
        If tCurBut(iCount).More Then BltPic tStyle.More, .hdc, .Width / 2 - 2, rHold.Bottom + 3, &HFFFFFF
    End With
Next iCount
'\\Buttons//
End Sub
'\\Public Functions//

'//Private Tools\\
Private Function FontExists(ByVal FontName As String) As Boolean 'Checks to see if a font is registered on the system
Dim iCount As Integer
For iCount = 0 To Screen.FontCount
    If (Screen.Fonts(iCount) = FontName) Then FontExists = True
Next iCount
End Function
Private Sub BltPic(ByVal Pic As IPictureDisp, ByVal DestDC As Long, ByVal x As Long, ByVal y As Long, Optional ByVal MaskColor As Long = vbNull) 'A function that will allow me to paint IDispPicture classes (Without the PaintPicture method). Its a generic function I've made awhile ago, so feel free to use it for anything.
Dim lHold As Long
lHold = CreateCompatibleDC(DestDC)
SelectObject lHold, Pic.Handle
If (MaskColor = vbNull) Then BitBlt DestDC, x, y, ScaleX(Pic.Width, vbHimetric, vbPixels), ScaleY(Pic.Height, vbHimetric, vbPixels), lHold, 0, 0, vbSrcCopy Else TransparentBlt DestDC, x, y, ScaleX(Pic.Width, vbHimetric, vbPixels), ScaleY(Pic.Height, vbHimetric, vbPixels), lHold, 0, 0, ScaleX(Pic.Width, vbHimetric, vbPixels), ScaleY(Pic.Height, vbHimetric, vbPixels), MaskColor
DeleteDC lHold
End Sub
Private Function InArray(ByVal Find As Long, Values() As Long) As Boolean
Dim iCount As Integer
For iCount = 0 To UBound(Values)
    If (Values(iCount) = Find) Then InArray = True: Exit Function
Next iCount
End Function
'\\Private Tools//
