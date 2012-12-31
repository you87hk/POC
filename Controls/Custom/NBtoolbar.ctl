VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl NBtoolbar 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11670
   ScaleHeight     =   420
   ScaleWidth      =   11670
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E-Open"
            Object.ToolTipText     =   "Open (F6)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E-New"
            Object.ToolTipText     =   "Add (F2)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E-Edit"
            Object.ToolTipText     =   "Edit (F5)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E-Delete"
            Object.ToolTipText     =   "Delete (F3)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E-Revise"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E-Copy"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E-Save"
            Object.ToolTipText     =   "Save (F10)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E-Cancel"
            Object.ToolTipText     =   "Cancel (F11)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E-Refresh"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E-Print"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E-Customer"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E-Vendor"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E-Item"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E-Exit"
            Object.ToolTipText     =   "Exit (F12)"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   120
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":1606
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":1A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":1D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":21C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":2616
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":2930
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":2C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":309C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":3978
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":3CA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":3FBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":42D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":45F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":4910
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NBtoolbar.ctx":4C2C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "NBtoolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Key constants:
'Entry (En)
Private Const EnOpen = "E-Open"
Private Const EnNew = "E-New"
Private Const EnEdit = "E-Edit"
Private Const EnDelete = "E-Delete"
Private Const EnRevise = "E-Revise"
Private Const EnCopy = "E-Copy"

Private Const EnSave = "E-Save"
Private Const EnCancel = "E-Cancel"

Private Const EnRefresh = "E-Refresh"
Private Const EnPrint = "E-Print"

Private Const EnCusSrh = "E-Customer"
Private Const EnVdrSrh = "E-Vendor"
Private Const EnItmSrh = "E-Item"

Private Const EnExit = "E-Exit"


Public Event Click(SelectButton As String)


Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)

    RaiseEvent Click(Button.Key)
    
End Sub

Private Sub UserControl_Initialize()


    With tbrProcess
        .BorderStyle = ccFixedSingle
        .Style = tbrFlat
        .ButtonHeight = 264.1895
        .ButtonWidth = 276.095
        .Align = vbAlignNone
        .Enabled = True
        .Visible = True
        .Appearance = cc3D
        .AllowCustomize = False
    End With
    
    
End Sub

Public Property Let ButtonVisible(ByVal vKey As String, ByVal vEnabled As Boolean)
    
    tbrProcess.Buttons(vKey).Visible = vEnabled
    
End Property

Public Property Let ButtonEnabled(ByVal vKey As String, ByVal vEnabled As Boolean)
    
    tbrProcess.Buttons(vKey).Enabled = vEnabled
    
End Property

Public Property Get ButtonEnabled(ByVal vKey As String) As Boolean

    ButtonEnabled = tbrProcess.Buttons(vKey).Enabled
    
End Property


Public Property Let ButtonToolTip(ByVal vKey As String, ByVal vTip As String)
    
    tbrProcess.Buttons(vKey).ToolTipText = vTip
    
End Property


Public Sub CheckKey(ByVal KeyCode As Integer, ByVal Shift As Integer)
                              
                              
        If Shift = vbCtrlMask Then
           
           Select Case KeyCode
                Case vbKeyF4
                    If tbrProcess.Buttons(EnCusSrh).Enabled = True Then _
                        RaiseEvent Click(EnCusSrh)
                Case vbKeyF5
                    If tbrProcess.Buttons(EnVdrSrh).Enabled = True Then _
                        RaiseEvent Click(EnVdrSrh)
                Case vbKeyF6
                    If tbrProcess.Buttons(EnItmSrh).Enabled = True Then _
                        RaiseEvent Click(EnItmSrh)
                Case Else
                    Exit Sub
            End Select
            
        Else
        
            Select Case KeyCode
                Case vbKeyF6
                    If tbrProcess.Buttons(EnOpen).Enabled = True Then _
                        RaiseEvent Click(EnOpen)
                Case vbKeyF2
                    If tbrProcess.Buttons(EnNew).Enabled = True Then _
                        RaiseEvent Click(EnNew)
                Case vbKeyF5
                    If tbrProcess.Buttons(EnEdit).Enabled = True Then _
                        RaiseEvent Click(EnEdit)
                Case vbKeyF3
                    If tbrProcess.Buttons(EnDelete).Enabled = True Then _
                        RaiseEvent Click(EnDelete)
                Case vbKeyF10
                    If tbrProcess.Buttons(EnSave).Enabled = True Then _
                        RaiseEvent Click(EnSave)
                Case vbKeyF11
                    If tbrProcess.Buttons(EnCancel).Enabled = True Then _
                        RaiseEvent Click(EnCancel)
                Case vbKeyF9
                    If tbrProcess.Buttons(EnPrint).Enabled = True Then _
                        RaiseEvent Click(EnPrint)
                Case vbKeyF12
                    If tbrProcess.Buttons(EnExit).Enabled = True Then _
                        RaiseEvent Click(EnExit)
                Case vbKeyF7
                    If tbrProcess.Buttons(EnRefresh).Enabled = True Then _
                        RaiseEvent Click(EnRefresh)
                Case Else
                    Exit Sub
            End Select
            
        End If
End Sub


