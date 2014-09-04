VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   672
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSwapSides 
      Caption         =   "Swap Alignment"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton btnCycleStyles 
      Caption         =   "Cycle Styles"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin Ribbon.BHRibbon BHRibbon1 
      Align           =   1  'Align Top
      Height          =   1740
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   3069
      Style           =   3
      Tabs            =   ""
      Cats            =   ""
      Buts            =   ""
      CharSet         =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":05B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BD9
            Key             =   ""
            Object.Tag             =   "&HFFFFFF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":119F
            Key             =   ""
            Object.Tag             =   "&HFFFFFF"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1752
            Key             =   ""
            Object.Tag             =   "&HFFFFFF"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub BHRibbon1_ButClick(Index As Integer, Caption As String, Category As Integer)
MsgBox "You clicked " & Caption & "!"
End Sub

Private Sub BHRibbon1_DlgClick(Index As Integer)
MsgBox "Dialog!"
End Sub

Private Sub btnSwapSides_Click()
BHRibbon1.RightLeft = Not BHRibbon1.RightLeft
End Sub

Private Sub btnCycleStyles_Click()
BHRibbon1.Style = BHRibbon1.Style Mod 3 + 1
End Sub

Private Sub Form_Load()
With BHRibbon1
    .ImageList = ImageList1
    .AddTab "Tab the First"
    .AddTab "Tab the Second"
    .AddCat 0, "Category 1"
    .AddCat 0, "Category 2", True
    .AddCat 0, "Cateogry 3"
    .AddCat 1, "Category 4"
    .AddCat 1, "Category 5"
    .AddBut 0, 0, 5, "Button 1", "The first button", True
    .AddBut 0, 0, 0, "Button 2", "The second button"
    .AddBut 0, 0, 2, "Button 3", "The third button"
    .AddBut 0, 1, 1, "Button 4", "The fourth button"
    .AddBut 1, 0, 1, "Button 5", "The fifth button"
    .AddBut 1, 0, 1, "Button 6", "The sixth button"
    .Refresh
End With
End Sub
