VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_DemandeAdoptionFiche 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demandes d'adoptions"
   ClientHeight    =   8760
   ClientLeft      =   1260
   ClientTop       =   1875
   ClientWidth     =   9510
   ControlBox      =   0   'False
   Icon            =   "frm_DemandeAdoptionFiche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   8640
         Picture         =   "frm_DemandeAdoptionFiche.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   375
         Left            =   7920
         Picture         =   "frm_DemandeAdoptionFiche.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar stb_DemandesFiche 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   8505
      Width           =   9510
      _ExtentX        =   16775
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
            TextSave        =   "18/07/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10:45"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8670
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
   Begin VB.Frame Fra_demande 
      Caption         =   "Demande d'adoption"
      Height          =   7815
      Left            =   120
      TabIndex        =   24
      ToolTipText     =   "Nature de la demande"
      Top             =   600
      Width           =   9255
      Begin VB.CheckBox chkAdoptions 
         Caption         =   "À inclure dans la recherche des demandes à actualiser"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   21
         ToolTipText     =   "Demande à inclure dans la procédure de recherche des demandes à actualiser ?"
         Top             =   7320
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.TextBox Txt_rema 
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   18
         Top             =   6360
         Width           =   6255
      End
      Begin VB.ComboBox Cmb_NatureDemande 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Nature de la demande"
         Top             =   1680
         Width           =   2775
      End
      Begin VB.ComboBox Cmb_TypeDemande 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Type de la demande"
         Top             =   1320
         Width           =   2775
      End
      Begin VB.ComboBox Cmb_AssMat 
         Height          =   315
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Demande d'Assistante Maternelle : (O)ui / (N)on"
         Top             =   1320
         Width           =   495
      End
      Begin MSComCtl2.DTPicker DTP_DateSortie 
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         ToolTipText     =   "Date de sortie de la demande"
         Top             =   6840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58785795
         CurrentDate     =   36287
      End
      Begin VB.TextBox Txt_MotifSortie 
         Height          =   285
         Left            =   4560
         MaxLength       =   50
         TabIndex        =   20
         ToolTipText     =   "Motif de la sortie de la demande"
         Top             =   6840
         Width           =   3375
      End
      Begin VB.TextBox Txt_numDossier 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.Frame Fra_Identification 
         Caption         =   "Identification"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   9015
         Begin MSComCtl2.DTPicker DTP_DateDemande 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1036
               SubFormatType   =   3
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   4
            ToolTipText     =   "Date de la demande"
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyy"
            Format          =   58785795
            CurrentDate     =   2958465
         End
         Begin VB.TextBox Txt_PrenomDemandeur 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5880
            TabIndex        =   3
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox Txt_NomDemandeur 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            TabIndex        =   2
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox Txt_NumDemande 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   1
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Lbl_DateDemande 
            Caption         =   "Date de la demande"
            Height          =   255
            Left            =   4320
            TabIndex        =   31
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Lbl_NumDemande 
            Caption         =   "Numéro de demande"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lbl_NumDossier 
            Caption         =   "Numéro de dossier"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Fra_Souhaits 
         Caption         =   "Souhaits"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   9015
         Begin VB.TextBox Txt_NombreEnfants 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7080
            TabIndex        =   13
            ToolTipText     =   "Nombre d'enfants demandés"
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Data Dat_OriginesEthniques 
            Caption         =   "Dat_originesethniques"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   6000
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   240
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.ComboBox Cmb_OrigEt 
            Height          =   315
            Index           =   0
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   240
            Width           =   3015
         End
         Begin VB.ComboBox Cmb_OrigEt 
            Height          =   315
            Index           =   1
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   600
            Width           =   3015
         End
         Begin VB.ComboBox Cmb_OrigEt 
            Height          =   315
            Index           =   2
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   960
            Width           =   3015
         End
         Begin VB.ComboBox Cmb_OrigEt 
            Height          =   315
            Index           =   3
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1320
            Width           =   3015
         End
         Begin VB.Data Dat_Types 
            Caption         =   "Dat_Types"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   720
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Data Dat_Natures 
            Caption         =   "Dat_Natures"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   1200
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CommandButton cmd_SupprimerEnfant 
            Height          =   375
            Left            =   2400
            Picture         =   "frm_DemandeAdoptionFiche.frx":0ACE
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Supprimer l'enfnt souhaité sélectionné"
            Top             =   1920
            Width           =   375
         End
         Begin VB.CommandButton cmd_CreerEnfant 
            Height          =   375
            Left            =   1200
            Picture         =   "frm_DemandeAdoptionFiche.frx":0BD0
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Créer un nouveau souhait"
            Top             =   1920
            Width           =   375
         End
         Begin VB.CommandButton cmd_ModifierEnfant 
            Height          =   375
            Left            =   1800
            Picture         =   "frm_DemandeAdoptionFiche.frx":0CD2
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Modifier l'enfant souhaité sélectionné"
            Top             =   1920
            Width           =   375
         End
         Begin VB.ComboBox Cmb_Fratrie 
            Height          =   315
            Left            =   5040
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Demande  de fratrie : (O)ui / (N)on"
            Top             =   1920
            Width           =   495
         End
         Begin MSComctlLib.ListView lvwAdoptions 
            Height          =   1200
            Index           =   0
            Left            =   1200
            TabIndex        =   17
            Top             =   2400
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   2117
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Numéro"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Sexe"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Handicap"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Tranche d'âges"
               Object.Width           =   2469
            EndProperty
         End
         Begin VB.Label Lbl_NombreEnfants 
            Caption         =   "Nombre d'enfants"
            Height          =   255
            Left            =   5760
            TabIndex        =   36
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Lbl_origine 
            Caption         =   "Origine ethnique :"
            Height          =   255
            Left            =   1320
            TabIndex        =   39
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Lbl_Fratrie 
            Caption         =   "Fratrie"
            Height          =   255
            Left            =   4440
            TabIndex        =   35
            Top             =   1920
            Width           =   495
         End
      End
      Begin VB.Label Lbl_rema 
         Caption         =   "Remarques"
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   6360
         Width           =   975
      End
      Begin VB.Label Lbl_MotifSortie 
         Caption         =   "Motif de sortie"
         Height          =   255
         Left            =   3480
         TabIndex        =   38
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Label Lbl_DateSortie 
         Caption         =   "Dossier sorti le"
         Height          =   255
         Left            =   480
         TabIndex        =   37
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Label Lbl_DemandeAssistanteMaternelle 
         Caption         =   "Demande assistante maternelle"
         Height          =   255
         Left            =   5040
         TabIndex        =   34
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Lbl_NatureDemande 
         Caption         =   "Nature de la demande"
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Lbl_TypeDemande 
         Caption         =   "Type de demande"
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   1320
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm_DemandeAdoptionFiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mfrmDemandeAdoptionEnfant As frm_DemandeAdoptionEnfant
Attribute mfrmDemandeAdoptionEnfant.VB_VarHelpID = -1

Private Enum IndexCheckBox          'Index des contrôles CheckBox
    chkINCLURE_DANS_RECHERCHE = 0
End Enum
Private Enum IndexColumnHeader      'Index des contrôles ColumnHeader
    hdrNUMÉRO = 0
    hdrSEXE
    hdrHANDICAP
    hdrÂGE
End Enum
Private Enum IndexListView          'Index des contrôles ListView
    lvwENFANTS = 0
End Enum

Private Const mstrNOMBRE_ENFANTS As String = " enfant(s)"

Private mrstEnfantsSouhaités As Recordset

Private QuitterCroix As Boolean
Private ValidéParMAJEnfant As Boolean
Private AnnulerLaCreation As Boolean

Public Sub ListeOriginesEthniques(Boite As ComboBox)
'Place la liste des origines ethniques dans la combo
    With Dat_OriginesEthniques
        .RecordSource = "SELECT ORET_CODE, ORET_LIBELLE FROM ORET ORDER BY ORET_CODE"
        .Refresh
        Boite.Clear
        Boite.AddItem ("")
        While Not .Recordset.EOF
            Boite.AddItem (.Recordset.Fields("ORET_LIBELLE")) + " <" + CStr(.Recordset.Fields("ORET_CODE")) + ">"
            .Recordset.MoveNext
        Wend
    End With
End Sub

Private Sub chkAdoptions_GotFocus(Index As Integer)
Const pkstrNomProcédure As String = "chkAdoptions_GotFocus"
On Error GoTo Err_chkAdoptions_GotFocus

    If (stb_DemandesFiche.Panels(4).Text = "CREATION") Or (stb_DemandesFiche.Panels(4).Text = "MODIFICATION") Then stb_DemandesFiche.SimpleText = chkAdoptions(Index).ToolTipText

    Exit Sub
    
Err_chkAdoptions_GotFocus:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Cmb_AssMat_Gotfocus()
    If (stb_DemandesFiche.Panels(4).Text = "CREATION") Or (stb_DemandesFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DemandesFiche.SimpleText = "Choisissez s'il s'agit d'une demande d'assistante maternelle ou non"
    End If
End Sub

Private Sub Cmb_Fratrie_Gotfocus()
    If (stb_DemandesFiche.Panels(4).Text = "CREATION") Or (stb_DemandesFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DemandesFiche.SimpleText = "Choisissez si le souhait concerne une fratrie ou non"
    End If
End Sub

Private Sub Cmb_NatureDemande_Gotfocus()
    If (stb_DemandesFiche.Panels(4).Text = "CREATION") Or (stb_DemandesFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DemandesFiche.SimpleText = "Choisissez la nature de la demande"
    End If
End Sub

Private Sub Cmb_TypeDemande_Gotfocus()
    If (stb_DemandesFiche.Panels(4).Text = "CREATION") Or (stb_DemandesFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DemandesFiche.SimpleText = "Choisissez le type de la demande"
    End If
End Sub

Private Sub cmd_CreerEnfant_Click()
Const pkstrNomProcédure As String = "cmd_CreerEnfant_Click"
On Error GoTo Err_cmd_CreerEnfant_Click

Dim sql As String

    If Val(Txt_NombreEnfants.Text) > 0 And Cmb_Fratrie.Text = "N" Then
        Exit Sub
    Else
        If stb_DemandesFiche.Panels(4).Text <> "MODIFICATION" Then
            cmd_Valider_Click
        End If
        cmd_Valider.Enabled = True
    
        stb_DemandesFiche.Panels(4) = "CREATION"
        
'       Ouvrir la fenêtre permettant le traitement des enfants
        With mfrmDemandeAdoptionEnfant
            .Initialiser comCRÉATION, True, CLng(Txt_NumDemande)
            .Show vbModal
        End With
        
        If Val(Txt_NombreEnfants) = 0 Then
             cmd_Valider.Enabled = False
        End If
        
        stb_DemandesFiche.Panels(4) = "MODIFICATION"
        AnnulerLaCreation = True
        
    End If

    Exit Sub
    
Err_cmd_CreerEnfant_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_ModifierEnfant_Click()
Const pkstrNomProcédure As String = "cmd_ModifierEnfant_Click"
Dim pitmEnfant As ListItem
On Error GoTo Err_cmd_ModifierEnfant_Click

'   Référencer l'enfant sélectionné
    Set pitmEnfant = lvwAdoptions(lvwENFANTS).SelectedItem
    
'   S'il existe
    If Not pitmEnfant Is Nothing Then
    
        stb_DemandesFiche.Panels(4) = "MODIFICATION"
        
'       Ouvrir la fenêtre permettant le traitement des enfants
        With mfrmDemandeAdoptionEnfant
            .Initialiser comMODIFICATION, True, CLng(Txt_NumDemande), pitmEnfant.Text
            .Show vbModal
        End With
        
    End If

    Exit Sub
    
Err_cmd_ModifierEnfant_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_Quitter_Click()

Dim requete As String
Dim sql As String
Dim RsFade As Recordset
Dim NumEnf As String

On Error GoTo Rupture_Integrite
    
  QuitterCroix = False
  If AnnulerLaCreation Then
    sql = "select ende_num from fade where dema_num=" + Iifent(Txt_NumDemande)
    Set RsFade = Db.OpenRecordset(sql)
    If Not RsFade.EOF Then
      RsFade.MoveFirst
      While Not RsFade.EOF
        NumEnf = RsFade.Fields(0)
        requete = "delete * from orso where ende_num=" + Iifent(NumEnf)
        Workspaces(0).Databases(0).Execute requete, dbFailOnError
        requete = "delete * from fade where ende_num=" + Iifent(NumEnf)
        Workspaces(0).Databases(0).Execute requete, dbFailOnError
        requete = "delete * from ende where ende_num=" + Iifent(NumEnf)
        Workspaces(0).Databases(0).Execute requete, dbFailOnError
        RsFade.Requery
      Wend
    End If
    
    requete = "DELETE * FROM DEMA WHERE DEMA_NUM = " + Iifent(Txt_NumDemande)
    Workspaces(0).Databases(0).Execute requete, dbFailOnError
  End If
  
  If InStr(Me.Caption, "Informations") = 0 Then
    If Val(Txt_NombreEnfants.Text) = 0 Then
      MsgBox "Il n'y a aucun enfant pour cette demande d'adoption !", vbCritical
      Exit Sub
    Else
      Unload Me
      frm_DossierGestion.Dat_demandes.Refresh
    End If
  Else
    Unload Me
  End If
    
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
  MousePointer = vbDefault
  Exit Sub
End Sub

Private Sub cmd_SupprimerEnfant_Click()
Const pkstrNomProcédure As String = "cmd_SupprimerEnfant_Click"
Dim pitmEnfant As ListItem
On Error GoTo Err_cmd_SupprimerEnfant_Click

    ValidéParMAJEnfant = True

'   Référencer l'enfant sélectionné
    Set pitmEnfant = lvwAdoptions(lvwENFANTS).SelectedItem
    
'   S'il existe
    If Not pitmEnfant Is Nothing Then
    
        stb_DemandesFiche.Panels(4) = "SUPPRESSION"
        
'       Ouvrir la fenêtre permettant le traitement des enfants
        With mfrmDemandeAdoptionEnfant
            .Initialiser comSUPPRESSION, True, CLng(Txt_NumDemande), pitmEnfant.Text
            .Show vbModal
        End With
    
        If Val(Txt_NombreEnfants) = 0 Then
             cmd_Valider.Enabled = False
        End If
            stb_DemandesFiche.Panels(4).Text = "MODIFICATION"
    
    End If

    Exit Sub
    
Err_cmd_SupprimerEnfant_Click:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub cmd_Valider_Click()
Const pkstrNomProcédure As String = "cmd_Valider_Click"
Dim pitmEnfant As ListItem
Dim requete As String
Dim sql As String
Dim Nbori, Nbenf
Dim RsCoac As Recordset
Dim RsFade As Recordset
Dim NumEnf As String
On Error GoTo Err_cmd_Valider_Click

QuitterCroix = False

If Not ValidéParMAJEnfant Then
    
  Select Case stb_DemandesFiche.Panels(4).Text
    Case "CREATION"
      MousePointer = vbHourglass
      requete = "INSERT INTO DEMA (DEMA_NUM, DOSS_NUM, DEMA_DATE, TYPE_CODE, NATU_CODE, DEMA_CDDAM, DEMA_CFRAT, DEMA_REMA, DEMA_InclureReactualisation)"
      requete = requete + "VALUES("
      requete = requete + Iifent(Txt_NumDemande) + ","
      requete = requete + Iifch(Txt_numDossier) + ","
      requete = requete + Iifch(DTP_DateDemande) + ","
      requete = requete + Iifent(CodeEntreCotes(Cmb_TypeDemande.Text)) + ","
      requete = requete + Iifent(CodeEntreCotes(Cmb_NatureDemande.Text)) + ","
      requete = requete + Iifch(Cmb_AssMat.Text) + ","
      requete = requete + Iifch(Cmb_Fratrie) + ","
      requete = requete + Iifch(Txt_rema) + ","
      requete = requete + IIf(chkAdoptions(chkINCLURE_DANS_RECHERCHE) = vbChecked, "True", "False") + ")"
      Workspaces(0).Databases(0).Execute requete, dbFailOnError
          
    Case "MODIFICATION"
      MousePointer = vbHourglass
      requete = "UPDATE DEMA"
      requete = requete + " SET DOSS_NUM=" + Iifch(Txt_numDossier)
      requete = requete + ",DEMA_DATE=" + Iifch(DTP_DateDemande)
      requete = requete + ",TYPE_CODE=" + Iifent(CodeEntreCotes(Cmb_TypeDemande.Text))
      requete = requete + ",NATU_CODE=" + Iifent(CodeEntreCotes(Cmb_NatureDemande.Text))
      requete = requete + ",DEMA_CDDAM=" + Iifch(Cmb_AssMat.Text)
      requete = requete + ",DEMA_CFRAT=" + Iifch(Cmb_Fratrie.Text)
      If DTP_DateSortie <> "" Then
          requete = requete + ",DEMA_DSORT=" + Iifch(DTP_DateSortie)
          requete = requete + ",DEMA_MSORT=" + Iifch(Txt_MotifSortie)
      Else
          requete = requete + ",DEMA_DSORT=" + "NULL"
          requete = requete + ",DEMA_MSORT="" """
      End If
      requete = requete + ",DEMA_REMA=" + Iifch(Txt_rema)
      requete = requete & ", DEMA_InclureReactualisation = " + IIf(chkAdoptions(chkINCLURE_DANS_RECHERCHE) = vbChecked, "True", "False")
      requete = requete + " WHERE DEMA_NUM=" + Iifent(Txt_NumDemande)
      Workspaces(0).Databases(0).Execute requete, dbFailOnError
      
'               Pour chaque enfant souhaité
                For Each pitmEnfant In lvwAdoptions(lvwENFANTS).ListItems
                
          sql = "DELETE * FROM ORSO WHERE ENDE_NUM=" + _
                 pitmEnfant.Text
          Workspaces(0).Databases(0).Execute sql, dbFailOnError
          
      For Nbori = 0 To 3
         If Cmb_OrigEt(Nbori).ListIndex >= 0 Then
                  sql = "INSERT INTO ORSO VALUES("
                  sql = sql + CodeEntreCotes(Cmb_OrigEt(Nbori).Text) + ","
                  sql = sql + pitmEnfant.Text + ")"
                  Workspaces(0).Databases(0).Execute sql, dbFailOnError
         End If
      Next
      
                Next pitmEnfant

      
      Unload Me
      frm_DossierGestion.Dat_demandes.Refresh

    Case "SUPPRESSION"
    
      MousePointer = vbHourglass
      sql = "select ende_num from coac where dema_num=" + Iifent(Txt_NumDemande)
      Set RsCoac = Db.OpenRecordset(sql)
      sql = "select ende_num from fade where dema_num=" + Iifent(Txt_NumDemande)
      Set RsFade = Db.OpenRecordset(sql)
      
      If Not RsCoac.EOF Then
        RsCoac.MoveFirst
        While Not RsCoac.EOF
          NumEnf = RsCoac.Fields(0)
          requete = "delete * from orso where ende_num=" + Iifent(NumEnf)
          Workspaces(0).Databases(0).Execute requete, dbFailOnError
          requete = "delete * from coac where ende_num=" + Iifent(NumEnf)
          Workspaces(0).Databases(0).Execute requete, dbFailOnError
          requete = "delete * from ende where ende_num=" + Iifent(NumEnf)
          Workspaces(0).Databases(0).Execute requete, dbFailOnError
          RsCoac.Requery
        Wend
      End If
      If Not RsFade.EOF Then
        RsFade.MoveFirst
        While Not RsFade.EOF
          NumEnf = RsFade.Fields(0)
          requete = "delete * from orso where ende_num=" + Iifent(NumEnf)
          Workspaces(0).Databases(0).Execute requete, dbFailOnError
          requete = "delete * from fade where ende_num=" + Iifent(NumEnf)
          Workspaces(0).Databases(0).Execute requete, dbFailOnError
          requete = "delete * from ende where ende_num=" + Iifent(NumEnf)
          Workspaces(0).Databases(0).Execute requete, dbFailOnError
          RsFade.Requery
        Wend
      End If
            
      requete = "DELETE * FROM SUPL WHERE PLAC_NUM IN ( "
      requete = requete + "SELECT PLAC_NUM FROM PLAC WHERE DEMA_NUM = " + Iifent(Txt_NumDemande) + ")"
      Workspaces(0).Databases(0).Execute requete, dbFailOnError
      requete = "DELETE * FROM ENPL WHERE PLAC_NUM IN ( "
      requete = requete + "SELECT PLAC_NUM FROM PLAC WHERE DEMA_NUM = " + Iifent(Txt_NumDemande) + ")"
      Workspaces(0).Databases(0).Execute requete, dbFailOnError
      requete = "DELETE * FROM PLAC WHERE DEMA_NUM = " + Iifent(Txt_NumDemande)
      Workspaces(0).Databases(0).Execute requete, dbFailOnError
      requete = "DELETE * FROM RENO WHERE DEMA_NUM =" + Iifent(Txt_NumDemande)
      Workspaces(0).Databases(0).Execute requete, dbFailOnError
      
      requete = "DELETE * FROM DEMA WHERE DEMA_NUM = " + Iifent(Txt_NumDemande)
      Workspaces(0).Databases(0).Execute requete, dbFailOnError
      Unload Me
      frm_DossierGestion.Dat_demandes.Refresh
    End Select
Else
    MousePointer = vbHourglass
    requete = "UPDATE DEMA"
    requete = requete + " SET DOSS_NUM=" + Iifch(Txt_numDossier)
    requete = requete + ",DEMA_DATE=" + Iifch(DTP_DateDemande)
    requete = requete + ",TYPE_CODE=" + Iifent(CodeEntreCotes(Cmb_TypeDemande.Text))
    requete = requete + ",NATU_CODE=" + Iifent(CodeEntreCotes(Cmb_NatureDemande.Text))
    requete = requete + ",DEMA_CDDAM=" + Iifch(Cmb_AssMat.Text)
    requete = requete + ",DEMA_CFRAT=" + Iifch(Cmb_Fratrie.Text)
    If DTP_DateSortie <> "" Then
        requete = requete + ",DEMA_DSORT=" + Iifch(DTP_DateSortie)
    End If
    requete = requete + ",DEMA_MSORT=" + Iifch(Txt_MotifSortie)
    requete = requete + ",DEMA_REMA=" + Iifch(Txt_rema)
    requete = requete & ", DEMA_InclureReactualisation = " + IIf(chkAdoptions(chkINCLURE_DANS_RECHERCHE) = vbChecked, "True", "False")
    requete = requete + " WHERE DEMA_NUM=" + Iifent(Txt_NumDemande)
    Workspaces(0).Databases(0).Execute requete, dbFailOnError
    Unload Me
    frm_DossierGestion.Dat_demandes.Refresh
End If
MousePointer = vbDefault

    Exit Sub
    
Err_cmd_Valider_Click:
Dim pbooAncienneGestionErreurs As Boolean
Dim pintTraitement As MessagesActions
'Affiche le code d'erreur approprié
    Select Case Err
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case 13:            pbooAncienneGestionErreurs = True: MsgBox "Valeur incompatible avec le type de colonne", vbOKOnly + vbExclamation
        Case 3022:          pbooAncienneGestionErreurs = True: MsgBox "Doublon" + ", création impossible", vbOKOnly + vbExclamation
        Case 3200:          pbooAncienneGestionErreurs = True: MsgBox "Reference" + ", suppression impossible", vbOKOnly + vbExclamation
        Case 3075:          pbooAncienneGestionErreurs = True: MsgBox "L'un des champs contient une cote, mise à jour impossible", vbOKOnly + vbExclamation
        Case 3315:          pbooAncienneGestionErreurs = True: MsgBox "Un champ obligatoire n'est pas renseigné, mise à jour impossible", vbOKOnly + vbExclamation
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
'   S'il s'agit de l'ancienne gestion d'erreurs
    If pbooAncienneGestionErreurs Then
        MousePointer = vbDefault
        Exit Sub
'   Sinon
    Else
        If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    End If

End Sub

Private Sub DTP_DateDemande_GotFocus()
    If (stb_DemandesFiche.Panels(4).Text = "CREATION") Or (stb_DemandesFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DemandesFiche.SimpleText = "Entrez la date de dépôt de la demande"
    End If
End Sub

Private Sub DTP_DateSortie_GotFocus()
    If (stb_DemandesFiche.Panels(4).Text = "CREATION") Or (stb_DemandesFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DemandesFiche.SimpleText = "Entrez la date de sortie de la demande"
    End If
End Sub

Private Sub Form_Load()

Dim i As Integer

    QuitterCroix = True
    ValidéParMAJEnfant = False
    AnnulerLaCreation = False
    
    Dat_OriginesEthniques.DatabaseName = CheminBase
    Dat_Natures.DatabaseName = CheminBase
    Dat_Types.DatabaseName = CheminBase
    
    ListeDemandeAssMat Cmb_AssMat
    Cmb_AssMat.ListIndex = 0
    ListeFratrie Cmb_Fratrie
    Cmb_Fratrie.ListIndex = 0
    ListeNaturesDemandes Cmb_NatureDemande
    Cmb_NatureDemande.ListIndex = 0
    ListeTypesDemandes Cmb_TypeDemande
    Cmb_TypeDemande.ListIndex = 0
    
    For i = 0 To 3
        ListeOriginesEthniques Cmb_OrigEt(i)
        Cmb_OrigEt(i).ListIndex = -1
    Next
    
    Select Case frm_DossierGestion.stb_GestionDossier.Panels(4)
        Case "CREER DEMANDE"
            stb_DemandesFiche.Panels(4).Text = "CREATION"
            Txt_MotifSortie.Enabled = False
            DTP_DateSortie.Enabled = False
        
        Case "MODIF DEMANDE"
            stb_DemandesFiche.Panels(4).Text = "MODIFICATION"
            
        Case "SUPPR DEMANDE"
            stb_DemandesFiche.Panels(4).Text = "SUPPRESSION"
    End Select

End Sub


Public Sub ListeDemandeAssMat(Boite As ComboBox)
'Permet une sélection OUI/NON gràce à la combo
    Boite.AddItem ("N")
    Boite.AddItem ("O")
End Sub

Public Sub ListeFratrie(Boite As ComboBox)
'Permet la sélection d'une fratrie gràce à la combo
    Boite.AddItem ("N")
    Boite.AddItem ("O")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const pkstrNomProcédure As String = "Form_Unload"
On Error GoTo Err_Form_Unload

    Set mrstEnfantsSouhaités = Nothing
    Set mfrmDemandeAdoptionEnfant = Nothing
    
    Exit Sub
    
Err_Form_Unload:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        

End Sub

Private Sub mfrmDemandeAdoptionEnfant_Validation(plngTraitement As CommunTraitements, pintNuméroEnfant As Integer)
Const pkstrNomProcédure As String = "mfrmDemandeAdoptionEnfant_Validation"
Dim pitmEnfant As ListItem
On Error GoTo Err_mfrmDemandeAdoptionEnfant_Validation

'   Implémenter la liste des enfants souhaités
    Implémenter "Enfants"
    
'   S'il ne s'agit pas d'une suppression
    If Not (plngTraitement = comSUPPRESSION Or pintNuméroEnfant = CInt(daoENREGISTREMENT_SUPPRIMÉ)) Then
'       Référencer l'enfant sur lequel a porté l'opération
        Set pitmEnfant = lvwAdoptions(lvwENFANTS).ListItems("*" & CStr(pintNuméroEnfant))
        pitmEnfant.Selected = True
    End If
    
    Exit Sub
    
Err_mfrmDemandeAdoptionEnfant_Validation:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

Private Sub Txt_MotifSortie_Gotfocus()
    If (stb_DemandesFiche.Panels(4).Text = "CREATION") Or (stb_DemandesFiche.Panels(4).Text = "MODIFICATION") Then
        stb_DemandesFiche.SimpleText = "Entrez le motif de sortie de la demande"
    End If
End Sub

Public Sub ListeNaturesDemandes(Boite As ComboBox)
'Place la liste des natures de demandes dans la combo
    With Dat_Natures
        .RecordSource = "SELECT NATU_LIBELLE, NATU_CODE FROM NATU ORDER BY NATU_LIBELLE"
        .Refresh
        Boite.Clear
        While Not Dat_Natures.Recordset.EOF
            Boite.AddItem (.Recordset.Fields("NATU_LIBELLE") + " <" + CStr(.Recordset.Fields("NATU_CODE")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub

Public Sub ListeTypesDemandes(Boite As ComboBox)
'Place la liste des types de demandes dans la combo
    With Dat_Types
        .RecordSource = "SELECT TYPE_LIBELLE, TYPE_CODE FROM TYPE ORDER BY TYPE_LIBELLE"
        .Refresh
        Boite.Clear
        While Not Dat_Types.Recordset.EOF
            Boite.AddItem (.Recordset.Fields("TYPE_LIBELLE") + " <" + CStr(.Recordset.Fields("TYPE_CODE")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub


'Implémenter les objets
'   pstrListeObjets:        liste des objets à implémenter au format '<objet#1>[|<objet#2>[...|<objet#n]]'
'   pstrListeParamètres:    liste des paramètres au format '<paramètre#1 de l'objet#1[¤paramètre#2 de l'objet#1[...¤paramètre#n de l'objet#1]]>[|<paramètre#1 de l'objet#2[¤paramètre#2 de l'objet#2[...¤paramètre#n de l'objet#2]]>[...|<paramètre#1 de l'objet#n[¤paramètre#2 de l'objet#n[...¤paramètre#n de l'objet#n]]>]]'

Private Sub Implémenter(pstrListeObjets As String, Optional pstrListeParamètres As String = "")
Const pkstrNomProcédure As String = "Implémenter"
Const pkintNom As Integer = 0, pkintIndex As Integer = 1
Dim mrstEnfantsSouhaités As Recordset
Dim pitmEnfant As ListItem
Dim pstrListesObjets() As String, pstrListesParamètres() As String, pstrParamètres() As String
Dim pintItérationObjet As Integer, pintItérationOrigine As Integer
On Error GoTo Err_Implémenter
    
    pstrListesObjets() = Split(pstrListeObjets, ","): pstrListesParamètres() = Split(pstrListeParamètres, gstrSÉPARATEUR_1)
    
'   Pour chaque objet dont on souhaite l'implémentation
    For pintItérationObjet = 0 To UBound(pstrListesObjets)
    
        Select Case pstrListesObjets(pintItérationObjet)
        
'           Enfants souhaités
            Case "Enfants"
        
                lvwAdoptions(lvwENFANTS).ListItems.Clear
                
'               Si les enfants souhaités n'ont pas encore été lus
                If mrstEnfantsSouhaités Is Nothing Then
'                   Les lire
                    Set mrstEnfantsSouhaités = modDAO.rstSQLExtraction(daoENFANTS_DEMANDÉS_OU_ACCORDÉS, True, gstrCONSULTATION_LIMITÉE, "FADE", CLng(Txt_NumDemande))
                    If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then
                        modCommun.DéfinirVerrouillageContrôle cmd_Valider, True
                        Exit Sub
                    End If
'               Sinon
                Else
'                   L'actualiser
                    mrstEnfantsSouhaités.Requery
                End If
                
                With mrstEnfantsSouhaités
'                   Pour chacun
                    While Not .EOF
'                       Créer un item sans image
                        Set pitmEnfant = lvwAdoptions(lvwENFANTS).ListItems.Add(, "*" & CStr(![ENDE_NUM]), CStr(![ENDE_NUM]))
                        modCommun.ImplémenterSubItems pitmEnfant, modCommun.varValeurFormatéeSiNull(![ENDE_SEXE], vbString), modCommun.varValeurFormatéeSiNull(![ENDE_HAND], vbString), modCommun.varValeurFormatéeSiNull(![ENDE_TRAGE], vbString)
                        .MoveNext
                    Wend
                    Txt_NombreEnfants = CStr(.RecordCount) & mstrNOMBRE_ENFANTS
               End With
        
        End Select
        
    Next pintItérationObjet
    
    Exit Sub
    
Err_Implémenter:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub


'Initialiser la fenêtre
'   plngTraitement:     type de traitement
'   pstrNuméroDossier:  numéro de dossier
'   pvarNuméroDemande:  numéro de demande [obligatoire s'il ne s'agit pas d'une création]

Public Sub Initialiser(plngTraitement As CommunTraitements, pstrNuméroDossier As String, Optional pvarNuméroDemande As Variant)
Const pkstrNomProcédure As String = "Initialiser"
Dim prstDemandeur As Recordset, prstOriginesEthniques As Recordset, prstDemande As Recordset, prstCount As Recordset
Dim pitmEnfant As ListItem
Dim pstrTitre As String
Dim pintNuméroEnfant As Integer, pintItération As Integer
Dim pbooCréationOuModification As Boolean, pbooCréationOuConsultation As Boolean
On Error GoTo Err_Initialiser
    
    Screen.MousePointer = vbHourglass
    
'   Centrer la feuille
    modAdoptions.CentrerFeuille Me
        
    Set mfrmDemandeAdoptionEnfant = New frm_DemandeAdoptionEnfant
    
'   Positionner les numéros de dossier et de demande
    Txt_numDossier = pstrNuméroDossier
    Txt_NumDemande = IIf(plngTraitement = comCRÉATION, NouveauCode("DEMA", "DEMA_NUM"), CStr(pvarNuméroDemande))
    
'   Lire le demandeur
    Set prstDemandeur = modDAO.rstSQLExtraction(daoINDIVIDU, False, gstrMESSAGE_ARRÊT_SUR_VERROUILLAGE & gstrSÉPARATEUR_1 & "Le demandeur n'existe plus.", pstrNuméroDossier, gstrTYPE_DEMANDEUR)
    If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then End
    With prstDemandeur
        If Not .RecordCount = 0 Then
            Txt_NomDemandeur = modCommun.varValeurFormatéeSiNull(![INDI_NOM], vbString)
            Txt_PrenomDemandeur = modCommun.varValeurFormatéeSiNull(![INDI_PREN], vbString)
        End If
    End With
                       
    Select Case plngTraitement
        Case comCRÉATION:       pstrTitre = "Création de":      pbooCréationOuModification = True: pbooCréationOuConsultation = True
        Case comCONSULTATION:   pstrTitre = "Informations sur": pbooCréationOuConsultation = True
        Case comMODIFICATION:   pstrTitre = "Modification de":  pbooCréationOuModification = True
        Case comSUPPRESSION:    pstrTitre = "Suppression de"
    End Select
    
'   Implémenter la liste des enfants souhaités
    Implémenter "Enfants"
        
'   S'il s'agit d'une création
    If plngTraitement = comCRÉATION Then
    
        Txt_NombreEnfants.Visible = False
        Lbl_NombreEnfants.Visible = False
        
        modCommun.PositionnerDateJour DTP_DateSortie, DTP_DateDemande
        DTP_DateSortie = ""
        DTP_DateDemande.Enabled = True
            
'   Sinon
    Else
    
'       Référencer l'enfant sélectionné
        Set pitmEnfant = lvwAdoptions(lvwENFANTS).SelectedItem
        If pitmEnfant Is Nothing Then pintNuméroEnfant = 0 Else pintNuméroEnfant = pitmEnfant.Text
        
'       Lire les origines ethniques souhaitées pour l'enfant
        Set prstOriginesEthniques = modDAO.rstSQLExtraction(daoEXTRACTION_003, False, gstrMESSAGE_ARRÊT_SUR_VERROUILLAGE, pintNuméroEnfant)
        If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then End
        With prstOriginesEthniques
            pintItération = 0
            While Not .EOF
                Cmb_OrigEt(pintItération).ListIndex = PresentDansListBox(Cmb_OrigEt(pintItération), "<" + CStr(![ORET_CODE]) + ">")
                .MoveNext
                pintItération = pintItération + 1
            Wend
        End With
        
'       Lire la demande
        Set prstDemande = modDAO.rstSQLExtraction(daoEXTRACTION_004, False, gstrMESSAGE_ARRÊT_SUR_VERROUILLAGE & gstrSÉPARATEUR_1 & "La demande, son type ou sa nature, n'existe plus.", pvarNuméroDemande)
        If Not modDAO.lngDernièreErreurDAO() = daoAUCUNE Then End
        With prstDemande
            DTP_DateDemande = ![DEMA_DATE]
            Cmb_AssMat.ListIndex = PresentDansListBox(Cmb_AssMat, CStr(![DEMA_CDDAM]))
            Cmb_Fratrie.ListIndex = PresentDansListBox(Cmb_Fratrie, CStr(![DEMA_CFRAT]))
            Cmb_NatureDemande.ListIndex = PresentDansListBox(Cmb_NatureDemande, "<" + CStr(![NATU_CODE]) + ">")
            Cmb_TypeDemande.ListIndex = PresentDansListBox(Cmb_TypeDemande, "<" + CStr(![TYPE_CODE]) + ">")
            Txt_MotifSortie = modCommun.varValeurFormatéeSiNull(![DEMA_MSORT], vbString)
            DTP_DateSortie.Value = modCommun.varValeurFormatéeSiNull(![DEMA_DSORT], vbDate, Date)
            If IsNull(![DEMA_DSORT]) Then DTP_DateSortie = ""
            Txt_rema = modCommun.varValeurFormatéeSiNull(![DEMA_REMA], vbString)
            chkAdoptions(chkINCLURE_DANS_RECHERCHE) = IIf(![DEMA_InclureReactualisation], vbChecked, vbUnchecked)
        End With
                   
    End If
    
    Txt_rema.Enabled = pbooCréationOuModification
    Txt_MotifSortie.Enabled = pbooCréationOuModification
    Cmb_AssMat.Enabled = pbooCréationOuModification
    Cmb_Fratrie.Enabled = pbooCréationOuModification
    Cmb_NatureDemande.Enabled = pbooCréationOuModification
    Cmb_TypeDemande.Enabled = pbooCréationOuModification
    DTP_DateSortie.Enabled = pbooCréationOuModification
    cmd_CreerEnfant.Visible = pbooCréationOuModification
    chkAdoptions(chkINCLURE_DANS_RECHERCHE).Enabled = pbooCréationOuModification
    cmd_ModifierEnfant.Visible = (plngTraitement = comMODIFICATION)
    cmd_SupprimerEnfant.Visible = (plngTraitement = comMODIFICATION)
    cmd_Valider.Enabled = Not pbooCréationOuConsultation
    For pintItération = 0 To 3
        Cmb_OrigEt(pintItération).Enabled = pbooCréationOuModification
    Next
            
    Me.Caption = pstrTitre & " la DEMANDE " & Txt_NumDemande & " / DOSSIER " + pstrNuméroDossier
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Err_Initialiser:
Dim pintTraitement As MessagesActions
    Select Case Err.Number
        Case vbObjectError: pintTraitement = mesINTERROMPRE
        Case Else:          pintTraitement = modMessages.lngAction(Err, pkstrNomProcédure, Me.Name, Nothing)
    End Select
    If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
        
End Sub

