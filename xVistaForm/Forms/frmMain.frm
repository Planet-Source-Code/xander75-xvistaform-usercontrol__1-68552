VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "xVistaForm"
   ClientHeight    =   5880
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   9855
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   7  'Size N S
   Picture         =   "frmMain.frx":058A
   ScaleHeight     =   5880
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xVistaFormProject.xVistaForm xVistaForm1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   688
      DisplayIcon     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      ForeColor       =   16777215
      FontItalic      =   0   'False
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Icon            =   "frmMain.frx":CE14
      MinHeight       =   3000
      MinWidth        =   4000
      ShowSytemTrayIcon=   -1  'True
      Style           =   1
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Show Close Button"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Show Maximise Button"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Show Minimise Button"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Display Icon"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton cmdStyle 
      Caption         =   "Change Form Style"
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkOptions_Click(Index As Integer)
    Select Case Index
        Case 0
            xVistaForm1.DisplayIcon = chkOptions(0).Value
        Case 1
            xVistaForm1.ShowMinimiseButton = chkOptions(1).Value
        Case 2
            xVistaForm1.ShowMaximiseButton = chkOptions(2).Value
        Case 3
            xVistaForm1.ShowCloseButton = chkOptions(3).Value
    End Select
End Sub

Private Sub cmdStyle_Click()
    If xVistaForm1.Style <> VistaDark Then
        xVistaForm1.Style = VistaDark
    Else
        xVistaForm1.Style = VistaLite
    End If
End Sub

Private Sub Form_Load()
    ' Create a system menu
    xVistaForm1.AddSysTrayItem 1, "Hide xVistaForm"
    xVistaForm1.AddSysTrayItem 2, "Show xVistaForm", , True, xGrayed
    xVistaForm1.AddSysTrayItem 3, "-"
    xVistaForm1.AddSysTrayItem 4, "Exit", True
End Sub

Private Sub xVistaForm1_Execute(ByVal ID As Long)
    ' ID corresponds to the ID given at the menu creation
    Select Case ID
        Case 1
            Me.Visible = False
            xVistaForm1.AmendSysTrayItem 1, , True, xGrayed
            xVistaForm1.AmendSysTrayItem 2, , False
        Case 2
            Me.Visible = True
            xVistaForm1.AmendSysTrayItem 1, , False
            xVistaForm1.AmendSysTrayItem 2, , True, xGrayed
        Case 4
            Unload Me
    End Select
End Sub
