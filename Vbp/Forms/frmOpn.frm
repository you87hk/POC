VERSION 5.00
Begin VB.Form frmOpn 
   BorderStyle     =   3  '���u�T�w��ܤ��
   ClientHeight    =   4185
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmOpn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame Frame1 
      Height          =   4170
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7200
      Begin VB.Image imgLogo 
         Height          =   945
         Left            =   360
         Picture         =   "frmOpn.frx":000C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  '�z��
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  '�a�k���
         BackStyle       =   0  '�z��
         Caption         =   "�L���t���U��"
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   3390
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  '�z��
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  '�a�k���
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6330
         TabIndex        =   5
         Top             =   2700
         Width           =   525
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  '�a�k���
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�����E��/�E�K"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4770
         TabIndex        =   6
         Top             =   2340
         Width           =   2085
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�i�P�s��X�t��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   27.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1680
         TabIndex        =   8
         Top             =   1560
         Width           =   4005
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "���v���s����ƨƷ~�������q"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�s����ƨƷ~�������q"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   18
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   7
         Top             =   960
         Width           =   3765
      End
      Begin VB.Image Image1 
         Height          =   4185
         Left            =   0
         Picture         =   "frmOpn.frx":1656
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7215
      End
   End
   Begin VB.Timer tmrUnload 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmOpn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = "���� " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = "�i�P�s��X�t��" 'App.Title
End Sub

Private Sub tmrUnload_Timer()
    Unload Me
End Sub
