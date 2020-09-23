VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Source Password"
   ClientHeight    =   2250
   ClientLeft      =   7800
   ClientTop       =   6975
   ClientWidth     =   4410
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   75
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   795
      Width           =   4275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Height          =   420
      Left            =   2955
      TabIndex        =   1
      Top             =   1440
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "If there is a database password, enter it below."
      Height          =   495
      Left            =   165
      TabIndex        =   2
      Top             =   150
      Width           =   3915
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

GetPassword = Text1.Text
Unload Me

End Sub
