VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Self Virus Check"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   $"frmMain.frx":000C
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":00ED
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please compile this Project before Running.
'Name the Exe file as SelfCheck.exe

Private Sub Command1_Click()
'Unlaod the Form.
Unload Me
'End the App.
End
End Sub

Private Sub Form_Load()
' Call the Function to check the EXE.
Call Virus_Check
End Sub
