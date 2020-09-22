VERSION 5.00
Object = "{D503CACF-B275-11D4-8CA6-4C9C0BC10000}#1.0#0"; "ZIconMaker.ocx"
Begin VB.Form frmTester 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmTester.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin ZIconMaker.ZIconM ZIconM1 
      Left            =   120
      Top             =   120
      _ExtentX        =   979
      _ExtentY        =   953
      GetPath         =   "E:\Archivos de programa\Microsoft Visual Studio\VB98\"
      SavePath        =   "E:\Archivos de programa\Microsoft Visual Studio\VB98\"
   End
   Begin VB.Label Label1 
      Caption         =   "Location"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Save to"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "frmTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************
'*************************************************
'Andres Zacarias
'Zacarias Software
'
'"Z Icon Maker test form"
'*************************************************
'*************************************************

Private Sub Command1_Click()
ZIconM1.GetPath = Text1.Text
ZIconM1.SavePath = Text2.Text

ZIconM1.MakeIt

End Sub

Private Sub Form_Load()

Text1.Text = App.Path & "\"
Text2.Text = App.Path & "\"

End Sub

