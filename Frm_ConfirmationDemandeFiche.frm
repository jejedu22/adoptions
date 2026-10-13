VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_ConfirmationDemandeFiche 
   Caption         =   "Fiche de confirmation de demande"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   ControlBox      =   0   'False
   Icon            =   "Frm_ConfirmationDemandeFiche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7170
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   6120
         Picture         =   "Frm_ConfirmationDemandeFiche.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   5520
         Picture         =   "Frm_ConfirmationDemandeFiche.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame Fra_Psy 
      Caption         =   "Rapport psychologue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3720
      TabIndex        =   5
      Top             =   2280
      Width           =   3015
      Begin MSComCtl2.DTPicker DTP_DemPsy 
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         ToolTipText     =   "Date de demande du rapport au psychologue"
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24641539
         CurrentDate     =   36404
      End
      Begin MSComCtl2.DTPicker DTP_RecPsy 
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         ToolTipText     =   "Date de réception du rapport du psychologue"
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24641539
         CurrentDate     =   36404
      End
      Begin VB.Label Label3 
         Caption         =   "Reçu le"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Demandé le"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Fra_as 
      Caption         =   "Rapport assistante sociale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   3015
      Begin MSComCtl2.DTPicker DTP_DemAs 
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         ToolTipText     =   "Date de demande du rapport à l'assistante sociale"
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24641539
         CurrentDate     =   36404
      End
      Begin MSComCtl2.DTPicker DTP_RecAs 
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         ToolTipText     =   "Date de réception du rapport de l'assistante sociale"
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24641539
         CurrentDate     =   36404
      End
      Begin VB.Label Label2 
         Caption         =   "Reçu le"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Lbl_dem 
         Caption         =   "Demandé le"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.TextBox Txt_annee 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   3
      ToolTipText     =   "N° d'ordre de l'année de la demande"
      Top             =   840
      Width           =   375
   End
   Begin VB.CheckBox Chk_DemReac 
      Caption         =   "Demande de réactualisation"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      ToolTipText     =   "Demande de réactualisation de la demande"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CheckBox Chk_RecMaintien 
      Caption         =   "Réception du maintien"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "Réception du courrier de maintien de la demande"
      Top             =   1200
      Width           =   1935
   End
   Begin MSComctlLib.StatusBar stb_InstitutionFiche 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   4275
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "18/01/00"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "16:45"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5971
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Lbl_année 
      Caption         =   "Année"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "Frm_ConfirmationDemandeFiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Quitter_Click()
  Unload Me
End Sub


Private Sub cmd_Valider_Click()

Dim sql, requete As String
Dim Rs As Recordset
Dim NbMois As Integer
Dim Newdate As Date

On Error GoTo Rupture_Integrite
  Workspaces(0).BeginTrans
  
  Select Case Frm_ConfirmationDemandeListe.stb_ConfDemListe.Panels(4).Text
    
    Case "CREATION"
      sql = "insert into reno values("
      sql = sql + Iifent(Txt_annee.Text) + ","
      sql = sql + Iifent(Frm_ConfirmationDemandeListe.Cmb_NumDemande.Text) + ","
      sql = sql + CStr(Chk_RecMaintien.Value) + ","
      sql = sql + CStr(Chk_DemReac.Value) + ","
      
      If DTP_DemAs <> "" Then
        sql = sql + Iifch(DTP_DemAs) + ","
        requete = "select * from para where para_code=5"
        Set Rs = Db.OpenRecordset(requete)
        NbMois = Rs("para_nbmois")
        Rs.Close
        Newdate = DateAdd("m", NbMois, DTP_DemAs)
        sql = sql + Iifch(CStr(Newdate)) + ","
      Else
        sql = sql + "Null" + "," + "Null" + ","
      End If
      If DTP_RecAs <> "" Then
        sql = sql + Iifch(DTP_RecAs) + ","
      Else
        sql = sql + "Null" + ","
      End If
      If DTP_DemPsy <> "" Then
        sql = sql + Iifch(DTP_DemPsy) + ","
        requete = "select * from para where para_code=5"
        Set Rs = Db.OpenRecordset(requete)
        NbMois = Rs("para_nbmois")
        Rs.Close
        Newdate = DateAdd("m", NbMois, DTP_DemPsy)
        sql = sql + Iifch(CStr(Newdate)) + ","
      Else
        sql = sql + "Null" + "," + "Null" + ","
      End If
      If DTP_RecPsy <> "" Then
        sql = sql + Iifch(DTP_RecPsy) + ")"
      Else
        sql = sql + "Null" + ")"
      End If
      
    Case "MODIFICATION"
      sql = "UPDATE RENO SET "
      sql = sql + "reno_recmait=" + CStr(Chk_RecMaintien.Value)
      sql = sql + ",reno_demrea=" + CStr(Chk_DemReac.Value)
      If DTP_DemAs <> "" Then
        sql = sql + ",reno_ddemmaas=" + Iifch(DTP_DemAs)
        requete = "select * from para where para_code=5"
        Set Rs = Db.OpenRecordset(requete)
        NbMois = Rs("para_nbmois")
        Rs.Close
        'Newdate = DateAdd("m", NbMois, DTP_DemRen)
        Newdate = DateAdd("m", NbMois, DTP_DemAs)
        sql = sql + ",reno_dprerecas=" + Iifch(CStr(Newdate))
      Else
        sql = sql + ",reno_ddemmaas=" + "Null"
        sql = sql + ",reno_dprerecas=" + "Null"
      End If
      If DTP_RecAs <> "" Then
        sql = sql + ",reno_drecmas=" + Iifch(DTP_RecAs)
      Else
        sql = sql + ",reno_drecmas=" + "Null"
      End If
      If DTP_DemPsy <> "" Then
        sql = sql + ",reno_ddemmps=" + Iifch(DTP_DemPsy)
        requete = "select * from para where para_code=5"
        Set Rs = Db.OpenRecordset(requete)
        NbMois = Rs("para_nbmois")
        Rs.Close
        Newdate = DateAdd("m", NbMois, DTP_DemPsy)
        sql = sql + ",reno_dprerecpsy=" + Iifch(CStr(Newdate))
      Else
        sql = sql + ",reno_ddemmps=" + "Null"
        sql = sql + ",reno_dprerecpsy=" + "Null"
      End If
      If DTP_RecPsy <> "" Then
        sql = sql + ",reno_drecmps=" + Iifch(DTP_RecPsy)
      Else
        sql = sql + ",reno_drecmps=" + "Null"
      End If
      sql = sql + " where conf_annee=" & Txt_annee
      sql = sql + " and dema_num=" & Frm_ConfirmationDemandeListe.Cmb_NumDemande
    
    Case "SUPPRESSION"
      sql = "delete * from reno where conf_annee=" & Txt_annee
      sql = sql + " and dema_num=" & Frm_ConfirmationDemandeListe.Cmb_NumDemande
  
  End Select
  
  Workspaces(0).Databases(0).Execute sql, dbFailOnError
  Workspaces(0).CommitTrans
  Unload Me
  Frm_ConfirmationDemandeListe.Dat_reno.Refresh
  Exit Sub
    
Rupture_Integrite:
'Affiche le code d'erreur approprié
    Select Case Err
        Case 13
            MsgBox "Valeur incompatible avec le type de colonne", vbOKOnly + vbExclamation
      
        Case 3022
            MsgBox "Doublon, création impossible", vbOKOnly + vbExclamation
      
        Case 3200
            MsgBox "Reference, suppression impossible", vbOKOnly + vbExclamation
      
        Case 3075
            MsgBox "L'un des champs contient une cote, mise à jour impossible", vbOKOnly + vbExclamation
      
        Case 3315
            MsgBox "Un champ obligatoire n'est pas renseigné, mise à jour impossible", vbOKOnly + vbExclamation
      
        Case Else
            MsgBox "Erreur" + CStr(Err) + " " + Err.Description + " survenue, mise à jour impossible", vbOKOnly + vbExclamation
    End Select
    Workspaces(0).Rollback
    Exit Sub
    
End Sub


Private Sub Form_Load()

  DTP_DemAs = Date
  DTP_DemPsy = Date
  DTP_RecAs = Date
  DTP_RecPsy = Date
  DTP_DemAs = ""
  DTP_DemPsy = ""
  DTP_RecAs = ""
  DTP_RecPsy = ""
   
End Sub


