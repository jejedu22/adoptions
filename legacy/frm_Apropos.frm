VERSION 5.00
Begin VB.Form frm_Apropos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A propos de ..."
   ClientHeight    =   4920
   ClientLeft      =   540
   ClientTop       =   1530
   ClientWidth     =   8475
   ControlBox      =   0   'False
   Icon            =   "frm_Apropos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txt_ProductName 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Txt_Productname"
      ToolTipText     =   "Nom de l'application"
      Top             =   180
      Width           =   5535
   End
   Begin VB.TextBox Txt_Version 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   4980
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Numéro de version de l'application"
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Txt_DateMaj 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   4980
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Date de dernière modification"
      Top             =   1380
      Width           =   3255
   End
   Begin VB.TextBox Txt_Localisation 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   4980
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Localisation du fichier exécutable"
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton cmd_Retour 
      Cancel          =   -1  'True
      Caption         =   "Retour au Menu de l'application"
      Default         =   -1  'True
      Height          =   495
      Left            =   2700
      TabIndex        =   0
      ToolTipText     =   "Retourner à l'application"
      Top             =   4260
      Width           =   5535
   End
   Begin VB.TextBox Txt_Comments 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00008000&
      Height          =   1455
      Left            =   2700
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frm_Apropos.frx":08CA
      Top             =   2580
      Width           =   5475
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4755
      Left            =   0
      Picture         =   "frm_Apropos.frx":08D0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Lab_Version 
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2700
      TabIndex        =   8
      Top             =   900
      Width           =   2115
   End
   Begin VB.Label Lab_DateMaj 
      Caption         =   "Dernière mise à jour le "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2700
      TabIndex        =   7
      Top             =   1440
      Width           =   2115
   End
   Begin VB.Label Lab_Localisation 
      Caption         =   "Localisation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2700
      TabIndex        =   5
      Top             =   1980
      Width           =   2115
   End
End
Attribute VB_Name = "frm_Apropos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Retour_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  Txt_ProductName.Text = App.ProductName
  Txt_Version = CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
  If Right(App.Path, 1) = "\" Then
    Txt_DateMaj = FileDateTime(App.Path + App.EXEName + ".EXE")
  Else
    Txt_DateMaj = FileDateTime(App.Path + "\" + App.EXEName + ".EXE")
  End If
  Txt_Localisation = App.Path
  Txt_Comments = App.Comments
End Sub
