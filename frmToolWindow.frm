VERSION 5.00
Begin VB.Form frmToolWindow 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "ToolBar"
   ClientHeight    =   75
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2715
   ClipControls    =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3.75
   ScaleMode       =   2  'Point
   ScaleWidth      =   135.75
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmToolWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Unload()

Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent Unload
End Sub
