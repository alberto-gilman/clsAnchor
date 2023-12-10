VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic 
      Height          =   1335
      Left            =   3960
      ScaleHeight     =   1275
      ScaleWidth      =   2835
      TabIndex        =   16
      Tag             =   "LBR"
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Frame FrameLEFTBOTTOMRIGHT 
      Caption         =   "LEFT BOTTOM"
      Height          =   1335
      Left            =   240
      TabIndex        =   12
      Tag             =   "LB"
      Top             =   6600
      Width           =   3615
      Begin VB.CommandButton cmdLEFTTOPRIGHTBOTTOM 
         Caption         =   "LEFT TOP RIGHT BOTTOM CLICK ME"
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Tag             =   "LTBR"
         Top             =   240
         Width           =   3375
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   480
      Tag             =   "LB"
      Top             =   5895
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   8130
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdLEFTBOTTOMRIGHT 
      Caption         =   "LEFT BOTTOM RIGHT"
      Height          =   615
      Left            =   2400
      TabIndex        =   6
      Tag             =   "LBR"
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton cmdRIGHTBOTTOM 
      Caption         =   "RIGHT BOTTOM"
      Height          =   615
      Left            =   4680
      TabIndex        =   5
      Tag             =   "RB"
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton cmdLEFTBOTTOM 
      Caption         =   "LEFT BOTTOM"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Tag             =   "LB"
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton cmdLEFTTOPRIGHT 
      Caption         =   "LEFT TOP RIGHT"
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Tag             =   "LTR, LEFT TOP RIGHT Button Clicked"
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdRIGHTTOP 
      Caption         =   "RIGHT TOP"
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Tag             =   "RT, RIGHT TOP Button Clicked"
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdLEFTTOP 
      Caption         =   "LEFT TOP"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Tag             =   "LT, LEFT TOP buton clicked"
      Top             =   840
      Width           =   2295
   End
   Begin TabDlg.SSTab SsTabFill 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Tag             =   "LTRB"
      Top             =   1560
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5741
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Fill"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdRIGHTTOPBOTTOM"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdLEFTTOPBOTTOM"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdLEFTTOPRIGHTBOTTON2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraRTB"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmdLEFTTOPRIGHTBOTTON2 
         Caption         =   "LEFT TOP RIGHT BOTTON"
         Height          =   2535
         Left            =   -74880
         TabIndex        =   17
         Tag             =   "LTRB"
         Top             =   480
         Width           =   5535
      End
      Begin VB.Frame fraRTB 
         Caption         =   "RIGHT TOP BOTTOM"
         Height          =   2415
         Left            =   240
         TabIndex        =   14
         Tag             =   "RTB"
         Top             =   600
         Width           =   6495
         Begin VB.ListBox LTBRList 
            Height          =   2010
            ItemData        =   "Form1.frx":0054
            Left            =   120
            List            =   "Form1.frx":005B
            TabIndex        =   18
            Tag             =   "LTBR"
            Top             =   240
            Width           =   4935
         End
         Begin VB.CommandButton cmdLEFTTOPRIGHTBOTTOM 
            Caption         =   "LEFT TOP RIGHT BOTTOM"
            Height          =   2055
            Index           =   1
            Left            =   5160
            TabIndex        =   15
            Tag             =   "LTRB"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "RIGHT TOP BOTTOM"
         Height          =   2535
         Left            =   -69240
         TabIndex        =   9
         Tag             =   "RTB"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdRIGHTTOPBOTTOM 
         Caption         =   "RIGHT TOP BOTTOM"
         Height          =   2535
         Left            =   -69240
         TabIndex        =   8
         Tag             =   "RTB"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdLEFTTOPBOTTOM 
         Caption         =   "LEFT TOP BOTTOM"
         Height          =   2535
         Left            =   -74880
         TabIndex        =   7
         Tag             =   "LTB"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Line Line 
         BorderWidth     =   10
         Tag             =   "LTRB"
         X1              =   -73560
         X2              =   -69360
         Y1              =   600
         Y2              =   2880
      End
   End
   Begin VB.Shape Shape 
      Height          =   495
      Index           =   0
      Left            =   240
      Tag             =   "LBR"
      Top             =   5880
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdLEFTTOP_Click()
    '"," is the value of Separator property of the Anchor instance we are using
    MsgBox Split(cmdLEFTTOP.Tag, ",")(1)
    
End Sub

Private Sub cmdLEFTTOPRIGHT_Click()
    '"," is the value of Separator property of the Anchor instance we are using
    MsgBox Split(cmdLEFTTOPRIGHT.Tag, ",")(1)

End Sub


Private Sub cmdLEFTTOPRIGHTBOTTOM_Click(Index As Integer)

    Dim F As Form2
    Set F = New Form2
    F.Show vbModal
End Sub

Private Sub cmdRIGHTTOP_Click()
    '"," is the value of Separator property of the Anchor instance we are using
    MsgBox Split(cmdRIGHTTOP.Tag, ",")(1)

End Sub


Private Sub Form_Resize()
    Static oAnchor As clsAnchor
    If Not oAnchor Is Nothing Then
    Else
        Set oAnchor = New clsAnchor
        'If no more properties are codified in Tag property the next sentence is not needed
        oAnchor.PropertyNumber = 0
        oAnchor.Form = Me
    End If
    
    oAnchor.Anchor
    
End Sub

