VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUnDock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'------------------------------------------------------------------------------
' Developed by Maksud Agamirzoyev 10/12/2000
'
' Class Properties: AutoDock (Boolean) - Sets whether dock will be
'                                        handled automatically. In other
'                                        words if try to undock and there
'                                        is already one undocked, it will
'                                        dock it back.
'
' Class Methods:    UnDock (Sub) - Undocks the controls from parent form
'                   Dock (Sub) - Docks control back to the parent form
'
'
' Note:
' It is not recomended to undock button controls, because then you might have some
' problems with receiving events for them, because of the way they handle messages.
'
' YOU SHOULD NOT HAVE CONTAINER OBJECT INSIDE OF ANOTHER CONTAINER OBJECT
'
' Request:
' Please if you modify this class send me a copy agamirm@westat.com
'-------------------------------------------------------------------------------
Option Explicit

'API calls
Private Declare Function GetWindowRect Lib "user32" _
                (ByVal hwnd As Long, lpRect As RECT) As Long
                
Private Declare Function GetParent Lib "user32" _
                (ByVal hwnd As Long) As Long

Private Declare Function GetWindow Lib "user32" _
                (ByVal hwnd As Long, ByVal wCmd As Long) As Long
                
Private Declare Function SetParent Lib "user32" _
                (ByVal hWndChild As Long, _
                ByVal hWndNewParent As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
                
Private Declare Function ShowWindow Lib "user32" _
                (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
                
Private Declare Function MoveWindow Lib "user32" _
                (ByVal hwnd As Long, ByVal x As Long, _
                ByVal y As Long, ByVal nWidth As Long, _
                ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias _
                "GetWindowLongA" (ByVal hwnd As Long, _
                ByVal nIndex As Long) As Long

'struct to store returned from GetWindowRect
Private Type RECT
    Left As Long
    Top As Long
    right As Long
    bottom As Long
    End Type

'status flags
Private Enum DockingStatus
    Docked = 0
    UnDocked = 1
    End Enum
    
Private Const SW_SHOWDEFAULT = 10
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Const WS_POPUP = &H80000000
Private Const WS_SYSMENU = &H80000
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDFIRST = 0
Private Const WS_GROUP = &H20000


'other local variables
Private hWndParent      As Long 'holds parent handle
Private hWndChild       As Long 'stores object handle

Private ObjRect         As RECT 'receives object size info
Private WinRect         As RECT 'receives parent size info

Private xOld            As Long 'holds old coordinates
Private yOld            As Long 'holds old coordinates

Private OldStyle            As Long
Private CurrentStatus       As DockingStatus 'stores current status of object
Private mvarAutoDock        As Boolean 'local copy
Private mvarSetAsChild      As Boolean 'local copy

Private WithEvents MyForm   As frmToolWindow 'Form which holds the undocked control
Attribute MyForm.VB_VarHelpID = -1

Public Property Let SetAsChild(ByVal vData As Boolean)
    'property which determines if Tool window is set to be
    'a child of controls parent
    mvarSetAsChild = vData
End Property

Public Property Get SetAsChild() As Boolean
    SetAsChild = mvarSetAsChild
End Property

Public Property Let AutoDock(ByVal vData As Boolean)
    'property which determines if auto dock should be allowed
    mvarAutoDock = vData
End Property

Public Property Get AutoDock() As Boolean
    AutoDock = mvarAutoDock
End Property

Public Sub UnDock(ByVal ObjectHandle As Long, Optional ProcessChildren As Boolean = False, Optional ByVal Caption As String, Optional xPos As Long = -1, Optional yPos As Long = -1)
    Dim RetVal     As Long
    Dim hTemp      As Long
    Dim ChildStyle As Long
        'check if control is undocked and autodock is off
        If CurrentStatus = UnDocked And AutoDock = False Then
           Exit Sub
        'check if control is undocked and autodock is on
        ElseIf CurrentStatus = UnDocked And AutoDock = True Then
           Dock 'dock control back to its parent
        End If
           
        hWndChild = ObjectHandle 'reassign to higher level variable
        RetVal = GetWindowRect(ObjectHandle, ObjRect) 'obtain size of object
        
        hWndParent = GetParent(ObjectHandle) 'get parent of object
        RetVal = GetWindowRect(hWndParent, WinRect) 'obtain size of object
        'checks if parent was obtained
        If hWndParent = 0 Then
           Err.Raise 10000, "UnDock", "No Parent"
           Exit Sub
        End If
        
        'check if position coorditates were specified,
        'if not assign from object actual postion
        If xPos = -1 And yPos = -1 Then
           'calculates postion of control relative to parent client area
           xPos = Abs(WinRect.Left - ObjRect.Left) - 4
           xOld = xPos 'saves in global variable
           yPos = Abs(WinRect.Top - ObjRect.Top) - 23
           yOld = yPos 'saves in global variable
        End If
        
        'saves old style information about control whic h is undocked
        OldStyle = GetWindowLong(ObjectHandle, GWL_STYLE)
        'changes flag of control moved to be a ChildWindow
        RetVal = SetWindowLong(ObjectHandle, GWL_STYLE, WS_CHILD)
        
        'checks if child processing requested, if control is frame for example
        If ProcessChildren = True Then
           hTemp = GetWindow(ObjectHandle, GW_CHILD) ' gets handle to the first child
           'enters loop to process all children
           Do Until hTemp = 0
              'changes flag of control moved to be a ChildWindow
              RetVal = SetWindowLong(hTemp, GWL_STYLE, WS_CHILD)
              'shows child window
              ShowWindow hTemp, SW_SHOWDEFAULT
              'gets next child window
              hTemp = GetWindow(hTemp, GW_HWNDNEXT)
           Loop
        End If
        
        'shows actual window on the new form
        ShowWindow ObjectHandle, SW_SHOWDEFAULT
        'rereferences parent of the control to ToolWindow
        SetParent ObjectHandle, MyForm.hwnd
        'rereferences parent of the control to be parent of ToolWindow
        If mvarSetAsChild = True Then _
           SetParent MyForm.hwnd, hWndParent
        
        'moves control
        MoveWindow ObjectHandle, 0, 0, (ObjRect.right - ObjRect.Left), _
                                 (ObjRect.bottom - ObjRect.Top), 1
        
        'moves toolwindow
        MoveWindow MyForm.hwnd, xPos, yPos, _
                   (ObjRect.right - ObjRect.Left) + 5, _
                   (ObjRect.bottom - ObjRect.Top) + 22, 1
        
        MyForm.Show 'shows form
        
        'assigns caption to the ToolWindow
        MyForm.Caption = Caption
        CurrentStatus = UnDocked 'records status that control was undocked
End Sub

Public Sub Dock()
        'checks if cotrol is undocked
        If CurrentStatus = UnDocked Then
           MyForm.Hide
           'sets cotrol back to its original parent
           SetParent hWndChild, hWndParent
           'restores original style settings
           SetWindowLong hWndChild, GWL_STYLE, OldStyle
           'positions it as it was
           MoveWindow hWndChild, xOld, yOld, _
                      (ObjRect.right - ObjRect.Left), _
                      (ObjRect.bottom - ObjRect.Top), 1
           'sets parent of ToolWindow to be Desktop
           SetParent MyForm.hwnd, 0
           CurrentStatus = Docked 'records that control is docked
        End If
End Sub

Private Sub Class_Initialize()
    Set MyForm = New frmToolWindow 'creates instance of the form
    'initializes variables
    CurrentStatus = Docked
    mvarAutoDock = False
    mvarSetAsChild = False
    Load MyForm 'loads the form into memory
End Sub

Private Sub MyForm_Unload()
    Dock 'docks control back when toolwindow is destroyed
End Sub
