VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ZIconM 
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1770
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1530
   ScaleWidth      =   1770
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Image imgIcon 
      Height          =   540
      Left            =   0
      Picture         =   "ZIconM.ctx":0000
      Top             =   0
      Width           =   555
   End
   Begin VB.Label etqGet 
      Caption         =   "Get"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Label etqSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "ZIconM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*************************************************
'*************************************************
'Andres Zacarias
'Zacarias Software
'
'"Z Icon Maker Project"
'*************************************************
'*************************************************
'Not for comercial use.


'This is the project and my little secret of making icons.
'Lets see if you can find it   :)


Option Explicit




'I decided to make this into an activex control becoz its more
'easier for me to put on projects.

'No need to to make an .ocx

Private Sub UserControl_Initialize()
  etqGet.Caption = (App.Path & "\")
  etqSave.Caption = (App.Path & "\")

End Sub

Private Sub UserControl_Resize()
  UserControl.Width = imgIcon.Width
  UserControl.Height = imgIcon.Height
End Sub

'This declares the new property
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("GetPath", etqGet.Caption, "")
    Call PropBag.WriteProperty("SavePath", etqSave.Caption, "")
End Sub

'This declares from were does etqcaption.Caption its going to read its its Caption
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    etqGet.Caption = PropBag.ReadProperty("GetPath", GetPath)
    etqSave.Caption = PropBag.ReadProperty("SavePath", SavePath)
End Sub

'This will get the Caption
Public Property Get GetPath() As String
    GetPath = etqGet.Caption
End Property

'This will get the Caption
Public Property Get SavePath() As String
    SavePath = etqSave.Caption
End Property

'And this will put the value on etqCaption.Caption
Public Property Let GetPath(ByVal NewGetPath As String)
    etqGet.Caption = NewGetPath
    PropertyChanged "GetPath"
End Property

'And this will put the value on etqCaption.Caption
Public Property Let SavePath(ByVal NewSavePath As String)
    etqSave.Caption = NewSavePath
    PropertyChanged "SavePath"
End Property


Public Function MakeIt()
'Dim Icon As String

    ' Load the picture into the ImageList.
    UserControl.ImageList1.ListImages.Add , , LoadPicture(UserControl.etqGet.Caption)
    
    ' Set the form's icon.
    'Set Icon = ImageList1.ListImages(1).ExtractIcon

    ' Save the icon file.
    SavePicture ImageList1.ListImages(1).ExtractIcon, UserControl.etqSave.Caption
End Function

