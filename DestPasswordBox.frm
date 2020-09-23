VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Destination Password"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Height          =   420
      Left            =   2925
      TabIndex        =   1
      Top             =   1530
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   45
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   900
      Width           =   4275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "If there is a database password, enter it below."
      Height          =   495
      Left            =   135
      TabIndex        =   2
      Top             =   255
      Width           =   3915
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
GetPassword = Text1.Text
Unload Me
End Sub
