VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   Caption         =   "Test"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   2295
      Begin VB.TextBox Text1 
         Height          =   1455
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Information"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   582
      ButtonWidth     =   609
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":08A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0CFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      CausesValidation=   0   'False
      Height          =   1620
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private obj As New DLLCtrUnDck.CUnDock

Private Sub Form_Load()
    'sets autodock to true, so when you try undock control
    'when other one is already undocked it will
    'automatically dock it
    obj.AutoDock = True
    
    'defines if ToolWindow will be set as child of undocked control parent window
    obj.SetAsChild = True
    
    'fills combobox
    With Me.Combo1
         .AddItem "Visual Basic"
         .AddItem "Visual C++"
         .AddItem "Visual J++"
    End With
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'here we undock Listbox
    If Button = 2 Then _
       obj.undock List1.hWnd, False, "List1"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2
            MsgBox "This project is to test DLLCtrUnDck." & vbCrLf & _
                   "    To Undock List1 click Right Mouse button." & vbCrLf & _
                   "    To Undock Toolbar1 double click on empty area of Toolbar." & vbCrLf & _
                   "    To Undock Frame on which Text1 located click last button on the Toolbar." & vbCrLf & _
                   "    To Add items from Combo1 to List1 click plus button on the ToolBar." & vbCrLf & _
                   "    To Clear List1 click third button on the Toolbar."
        Case 3
            If Me.Combo1.Text <> "" Then
               Me.List1.AddItem Me.Combo1.Text
               Me.Text1.Text = Me.Text1.Text & Me.Combo1.Text & " added." & vbCrLf
            End If
        Case 4
            Me.List1.Clear
            Me.Text1.Text = Me.Text1.Text & "List1 cleared." & vbCrLf
        Case 6
            obj.undock Me.Frame1.hWnd, True, "Frame1"
    End Select
End Sub

Private Sub Toolbar1_DblClick()
    'here we undock toolbar
    obj.undock Toolbar1.hWnd, False, "Toolbar"
End Sub
