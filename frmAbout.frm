VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ICQ ActiveX Control"
   ClientHeight    =   3855
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5790
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2660.79
   ScaleMode       =   0  'User
   ScaleWidth      =   5437.109
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox OCXPicture 
      AutoSize        =   -1  'True
      Height          =   1860
      Left            =   105
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   1800
      ScaleWidth      =   2880
      TabIndex        =   0
      Top             =   105
      Width           =   2940
   End
   Begin VB.TextBox MiscellaneousText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1347
      Left            =   128
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmAbout.frx":2F41
      Top             =   2100
      Width           =   5506
   End
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   105
      TabIndex        =   6
      Top             =   1995
      Width           =   5580
   End
   Begin VB.Label lblOthersCopyright 
      Caption         =   "ICQ, the ICQ logo, the flower device as well as other marks, are trademarks of ICQ Inc."
      Height          =   645
      Left            =   3150
      TabIndex        =   4
      Top             =   1365
      Width           =   2535
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright © 2000-2001 MediEvilz Software. All rights reserved."
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   3570
      Width           =   5580
   End
   Begin VB.Label OCXVersion 
      Height          =   225
      Left            =   3150
      TabIndex        =   2
      Top             =   315
      Width           =   2535
   End
   Begin VB.Label OCXTitle 
      Height          =   225
      Left            =   3150
      TabIndex        =   1
      Top             =   105
      Width           =   2535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  OCXTitle.Caption = App.FileDescription
  OCXVersion.Caption = "Version " + _
    Trim$(Str$(App.Major)) + "." + _
    Trim$(Str$(App.Minor)) + "." + _
    Trim$(Str$(App.Revision))
End Sub
