VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Individu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Individu"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   ControlBox      =   0   'False
   Icon            =   "frm_Individu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Cmb_AssuranceMaladie 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3600
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTP_DateDeces 
      Height          =   300
      Left            =   5040
      TabIndex        =   16
      Top             =   4080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd/MM/yyy"
      Format          =   58851331
      CurrentDate     =   36285
   End
   Begin MSComCtl2.DTPicker DTP_DateDivorce 
      Height          =   300
      Left            =   5040
      TabIndex        =   14
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd/MM/yyy"
      Format          =   58851331
      CurrentDate     =   36285
   End
   Begin MSComCtl2.DTPicker DTP_DateAccueil 
      Height          =   300
      Left            =   7920
      TabIndex        =   17
      Top             =   4080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd/MM/yyy"
      Format          =   58851331
      CurrentDate     =   36285
   End
   Begin MSComCtl2.DTPicker DTP_DateVeuvage 
      Height          =   300
      Left            =   7920
      TabIndex        =   15
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd/MM/yyy"
      Format          =   58851331
      CurrentDate     =   36285
   End
   Begin VB.Frame fra_SituationActive 
      Caption         =   "Situation Active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   120
      TabIndex        =   29
      Top             =   4440
      Width           =   9255
      Begin VB.Data Dat_Profession 
         Caption         =   "Dat_Profession"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data Dat_AssuranceMaladie 
         Caption         =   "Dat_AssuranceMaladie"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox Txt_TelEmployeur 
         Height          =   285
         Left            =   6840
         MaxLength       =   14
         TabIndex        =   23
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Txt_VilleEmployeur 
         Height          =   285
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   22
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Txt_CPEmployeur 
         Height          =   285
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   21
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Txt_Adr 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1080
         Width           =   6975
      End
      Begin VB.TextBox Txt_NomEmployeur 
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   19
         Top             =   720
         Width           =   3375
      End
      Begin VB.ComboBox Cmb_Profession 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label lbl_Profession 
         Caption         =   "Profession"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl_NomEmployeur 
         Caption         =   "Nom de l'employeur"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lbl_Adresse 
         Caption         =   "Adresse"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lbl_Telephone 
         Caption         =   "Téléphone"
         Height          =   255
         Left            =   6000
         TabIndex        =   43
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lbl_Ville 
         Caption         =   "Ville"
         Height          =   255
         Left            =   2040
         TabIndex        =   42
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lbl_CodePostal 
         Caption         =   "Code postal"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame fra_Outils 
      Height          =   585
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmd_Quitter 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   8040
         Picture         =   "frm_Individu.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Annuler"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton cmd_Valider 
         Default         =   -1  'True
         Height          =   375
         Left            =   8640
         Picture         =   "frm_Individu.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Valider"
         Top             =   150
         Width           =   375
      End
   End
   Begin VB.Frame fra_EtatCivil 
      Caption         =   "Etat civil de l'individu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Left            =   120
      TabIndex        =   26
      Top             =   600
      Width           =   9255
      Begin VB.TextBox Txt_EMail 
         Height          =   285
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   54
         ToolTipText     =   "Adresse électronique"
         Top             =   2490
         Width           =   6855
      End
      Begin VB.TextBox txt_nationalite 
         Height          =   285
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   52
         ToolTipText     =   "Nationalité de l'individu"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Txt_VillePre 
         Height          =   285
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   12
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox Txt_CPpre 
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   11
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox Txt_Adr2Pre 
         Height          =   285
         Left            =   4800
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox Txt_Adr1Pre 
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox Txt_DepartementNaissance 
         Height          =   285
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   8
         ToolTipText     =   "Département de naissance"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Txt_VilleNaissance 
         Height          =   285
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   7
         ToolTipText     =   "Lieu de naissance"
         Top             =   1080
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTP_DateNaissance 
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyy"
         Format          =   58851331
         CurrentDate     =   36283
      End
      Begin VB.TextBox Txt_Prenom 
         Height          =   285
         Left            =   6720
         MaxLength       =   20
         TabIndex        =   5
         ToolTipText     =   "Prénom de l'individu"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Txt_Nom 
         Height          =   285
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "Nom de l'individu"
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox Cmb_Titre 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         ToolTipText     =   "Titre"
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox Cmb_CodeFratrie 
         Height          =   315
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Txt_NumIndividu 
         Height          =   285
         Left            =   3600
         TabIndex        =   1
         ToolTipText     =   "Numéro de l'individu"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Txt_NumDossier 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "Numéro du dossier"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lbl_AdressePrecedente 
         Caption         =   "Adresse Précédente (-1 an)"
         Height          =   615
         Left            =   120
         TabIndex        =   40
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lbl_Departement 
         Caption         =   "Dept"
         Height          =   255
         Left            =   6120
         TabIndex        =   39
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lbl_LieuNaissance 
         Caption         =   "à"
         Height          =   255
         Left            =   3360
         TabIndex        =   38
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label lbl_Nationalité 
         Caption         =   "Nationalité"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lbl_NéLe 
         Caption         =   "Né(e) le"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lbl_Prenom 
         Caption         =   "Prénom"
         Height          =   255
         Left            =   6120
         TabIndex        =   35
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lbl_Nom 
         Caption         =   "Nom"
         Height          =   255
         Left            =   3120
         TabIndex        =   34
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lbl_Titre 
         Caption         =   "Titre"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lbl_CodeFratrie 
         Caption         =   "Fratrie"
         Height          =   180
         Left            =   6120
         TabIndex        =   32
         Top             =   405
         Width           =   615
      End
      Begin VB.Label lbl_NumIndividu 
         Caption         =   "N° individu"
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl_NumDossier 
         Caption         =   "N° dossier"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar stb_Individu 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   6255
      Width           =   9465
      _ExtentX        =   16695
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
            TextSave        =   "13/05/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "14:09"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10478
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl_DateAccueil 
      Caption         =   "Date d'accueil"
      Height          =   255
      Left            =   6600
      TabIndex        =   51
      ToolTipText     =   "Permet la sélection d'une date d'accueil"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lbl_DateVeuvage 
      Caption         =   "Date du veuvage"
      Height          =   255
      Left            =   6600
      TabIndex        =   50
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lbl_DateDeces 
      Caption         =   "Date du décès"
      Height          =   255
      Left            =   3960
      TabIndex        =   49
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lbl_DateDivorce 
      Caption         =   "Date divorce"
      Height          =   255
      Left            =   3960
      TabIndex        =   48
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lbl_CouvertureSociale 
      Caption         =   "Couverture sociale"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   3600
      Width           =   1335
   End
End
Attribute VB_Name = "frm_Individu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private QuitterCroix As Boolean

Private Sub Cmb_AssuranceMaladie_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Choisissez le code d'assurance maladie de l'individu"
    End If
End Sub

Private Sub Cmb_CodeFratrie_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Choisissez le code de fratrie de l'individu"
    End If
End Sub

Private Sub Cmb_Nationalite_Change()

End Sub

Private Sub Text1_Change()

End Sub


Private Sub txt_nationalite_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Choisissez la nationalité de l'individu"
    End If
End Sub

Private Sub Cmb_Profession_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Choisissez la profession de l'individu"
    End If
End Sub

Private Sub Cmb_Titre_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Choisissez le titre de l'individu"
    End If
End Sub

Private Sub cmd_Quitter_Click()
    QuitterCroix = False
    Unload Me
End Sub

Private Sub cmd_Valider_Click()
Const pkstrNomProcédure As String = "cmd_Valider_Click"
Dim prstDemandeurs As Recordset
Dim pstrMessage As String
Dim requete As String
On Error GoTo Err_cmd_Valider_Click
    
'   S'il ne s'agit pas d'une suppression
    If Txt_Nom.Enabled Then
'       S'il s'agit de valider les références du demandeur
        If CodeEntreCotes(Cmb_CodeFratrie) = gstrTYPE_DEMANDEUR Then
'           Vérifier l'existence de demandeurs portant mêmes nom et prénom
            Set prstDemandeurs = modDAO.rstSQLExtraction(daoEXTRACTION_011, True, gstrVALIDATION_ANNULÉE, IIf(Txt_Nom = "", Null, Txt_Nom), IIf(Txt_Prenom = "", Null, Txt_Prenom), CLng(Txt_NumIndividu))
            If Not (modDAO.lngDernièreErreurDAO() = daoAUCUNE Or modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT) Then Exit Sub
'           S'il en existe
            If Not modDAO.lngDernièreErreurDAO() = daoAUCUN_ENREGISTREMENT Then
                With prstDemandeurs
'                   Pour chacun
                    While Not .EOF
'                       Ajouter au message les références du dossier
                        pstrMessage = pstrMessage & vbCrLf & "- " & CStr(![DOSS_NUM])
                        .MoveNext
                    Wend
                End With
'               En informer l'utilisateur
                Select Case modMessages.lngRéponseUtilisateur(mesQUESTION_004, pstrMessage)
'                   L'utilisateur valide la saisie
                    Case vbYes
'                   L'utilisateur renonce à valider
                    Case vbNo:  Txt_Nom.SetFocus: Exit Sub
                End Select
            End If
        End If
    End If
    
QuitterCroix = False
Workspaces(0).BeginTrans

  Select Case stb_Individu.Panels(4).Text
    Case "CREATION"
      requete = "INSERT INTO INDI VALUES ("
      requete = requete + Iifent(Txt_NumIndividu) + ","
      If Cmb_Profession.Text <> "" Then
          requete = requete + CodeEntreCotes(Cmb_Profession.Text) + ","
      Else
          requete = requete + "NULL" + ","
      End If
      If Cmb_AssuranceMaladie.Text <> "" Then
          requete = requete + CodeEntreCotes(Cmb_AssuranceMaladie.Text) + ","
      Else
          requete = requete + "NULL" + ","
      End If
      requete = requete + Iifch(Txt_NumDossier) + ","
      If Cmb_CodeFratrie.Text <> "" Then
          requete = requete + Iifch(CodeEntreCotes(Cmb_CodeFratrie.Text)) + ","
      Else
          requete = requete + "NULL" + ","
      End If
      requete = requete + Iifch(Cmb_Titre.Text) + ","
      requete = requete + Iifch(Txt_Nom) + ","
      requete = requete + Iifch(Txt_Prenom) + ","
      requete = requete + Iifch(DTP_DateNaissance) + ","
      requete = requete + Iifch(Txt_VilleNaissance) + ","
      requete = requete + Iifent(Txt_DepartementNaissance) + ","
      requete = requete + Iifch(txt_nationalite.Text) + ","
      requete = requete + Iifch(Txt_Adr1Pre) + ","
      requete = requete + Iifch(Txt_Adr2Pre) + ","
      requete = requete + Iifch(Txt_CPpre) + ","
      requete = requete + Iifch(Txt_VillePre) + ","
      If DTP_DateDivorce <> "" Then
          requete = requete + Iifch(DTP_DateDivorce) + ","
      Else
          requete = requete + "NULL" + ","
      End If
      If DTP_DateVeuvage <> "" Then
          requete = requete + Iifch(DTP_DateVeuvage) + ","
      Else
          requete = requete + "NULL" + ","
      End If
      If DTP_DateDeces <> "" Then
          requete = requete + Iifch(DTP_DateDeces) + ","
      Else
          requete = requete + "NULL" + ","
      End If
      If DTP_DateAccueil <> "" Then
          requete = requete + Iifch(DTP_DateAccueil) + ","
      Else
          requete = requete + "NULL" + ","
      End If
      requete = requete + Iifch(Txt_EMail) + ","
      requete = requete + Iifch(Txt_NomEmployeur) + ","
      requete = requete + Iifch(Txt_Adr) + ","
      requete = requete + Iifch(Txt_CPEmployeur) + ","
      requete = requete + Iifch(Txt_VilleEmployeur) + ","
      requete = requete + Iifch(Txt_TelEmployeur) + ")"
      
    Case "MODIFICATION"
      requete = "UPDATE INDI"
      If Cmb_Profession.Text <> "" Then
          requete = requete + " SET PROF_CODE=" + CodeEntreCotes(Cmb_Profession.Text)
      Else
          requete = requete + " SET PROF_CODE=" + "NULL"
      End If
      If Cmb_AssuranceMaladie.Text <> "" Then
          requete = requete + ",ASMA_CODE=" + CodeEntreCotes(Cmb_AssuranceMaladie.Text)
      Else
          requete = requete + ",ASMA_CODE=" + "NULL"
      End If
      requete = requete + ",DOSS_NUM=" + Iifent(Txt_NumDossier)
      If Cmb_CodeFratrie.Text <> "" Then
          requete = requete + ",INDI_CDFR=" + Iifch(CodeEntreCotes(Cmb_CodeFratrie.Text))
      Else
          requete = requete + ",INDI_CDFR=" + "NULL"
      End If
      requete = requete + ",INDI_TITR=" + Iifch(Cmb_Titre.Text)
      requete = requete + ",INDI_NOM=" + Iifch(Txt_Nom)
      requete = requete + ",INDI_PREN=" + Iifch(Txt_Prenom)
      requete = requete + ",INDI_DNAIS=" + Iifch(DTP_DateNaissance)
      requete = requete + ",INDI_LNAIS=" + Iifch(Txt_VilleNaissance)
      requete = requete + ",INDI_DPNAI=" + Iifent(Txt_DepartementNaissance)
      requete = requete + ",INDI_NATIO=" + Iifch(txt_nationalite)
      requete = requete + ",INDI_ADRPRE1=" + Iifch(Txt_Adr1Pre)
      requete = requete + ",INDI_ADRPRE2=" + Iifch(Txt_Adr2Pre)
      requete = requete + ",INDI_CODPRE=" + Iifch(Txt_CPpre)
      requete = requete + ",INDI_VILPRE=" + Iifch(Txt_VillePre)
      If DTP_DateDivorce <> "" Then
          requete = requete + ",INDI_DDIVO=" + Iifch(DTP_DateDivorce)
      Else
          requete = requete + ",INDI_DDIVO=" + "NULL"
      End If
      If DTP_DateVeuvage <> "" Then
          requete = requete + ",INDI_DVEUV=" + Iifch(DTP_DateVeuvage)
      Else
          requete = requete + ",INDI_DVEUV=" + "NULL"
      End If
      If DTP_DateDeces <> "" Then
          requete = requete + ",INDI_DDECE=" + Iifch(DTP_DateDeces)
      Else
          requete = requete + ",INDI_DDECE=" + "NULL"
      End If
      If DTP_DateAccueil <> "" Then
          requete = requete + ",INDI_DACCU=" + Iifch(DTP_DateAccueil)
      Else
          requete = requete + ",INDI_DACCU=" + "NULL"
      End If
      requete = requete + ",INDI_EMAIL=" + Iifch(Txt_EMail)
      requete = requete + ",TRAV_NOMEMP=" + Iifch(Txt_NomEmployeur)
      requete = requete + ",TRAV_ADDEMP=" + Iifch(Txt_Adr)
      requete = requete + ",TRAV_CPEMP=" + Iifch(Txt_CPEmployeur)
      requete = requete + ",TRAV_VILLEMP=" + Iifch(Txt_VilleEmployeur)
      requete = requete + ",TRAV_TELEMP=" + Iifch(Txt_TelEmployeur)
      
      requete = requete + " WHERE INDI_NUM=" + Iifent(Txt_NumIndividu)
      
    Case "SUPPRESSION"
      requete = "DELETE * FROM INDI WHERE INDI_NUM = " + Iifent(Txt_NumIndividu)
      
  End Select
  Workspaces(0).Databases(0).Execute requete, dbFailOnError
  Workspaces(0).CommitTrans
  Unload Me
  frm_DossierGestion.Dat_dossiers.Refresh
  
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
    Workspaces(0).Rollback
'   S'il s'agit de l'ancienne gestion d'erreurs
    If pbooAncienneGestionErreurs Then
        MousePointer = vbDefault
        Exit Sub
'   Sinon
    Else
        If pintTraitement = mesPOURSUIVRE Then Resume Next Else End
    End If

End Sub

Private Sub DTP_DateAccueil_GotFocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez la date d'accueil (s'il y a lieu)"
    End If
End Sub

Private Sub DTP_DateDeces_GotFocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez la date de décès (s'il y a lieu)"
    End If
End Sub

Private Sub DTP_DateDivorce_GotFocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez la date de divorce (s'il y a lieu)"
    End If
End Sub

Private Sub DTP_DateNaissance_GotFocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez la date de naissance de l'individu"
    End If
End Sub

Private Sub DTP_DateVeuvage_GotFocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez la date de veuvage (s'il y a lieu)"
    End If
End Sub

Private Sub Form_Load()
    QuitterCroix = True
    If frm_DossierGestion.stb_GestionDossier.Panels(4) = "CREATION INDIV" Then
        stb_Individu.Panels(4).Text = "CREATION"
    End If
    If frm_DossierGestion.stb_GestionDossier.Panels(4) = "MODIF. INDIVIDU" Then
        stb_Individu.Panels(4).Text = "MODIFICATION"
    End If
    If frm_DossierGestion.stb_GestionDossier.Panels(4) = "SUPPR. INDIVIDU" Then
        stb_Individu.Panels(4).Text = "SUPPRESSION"
    End If

    Dat_AssuranceMaladie.DatabaseName = CheminBase
    Dat_Profession.DatabaseName = CheminBase
    ListeAssuranceMaladie Cmb_AssuranceMaladie
    ListeProfessions Cmb_Profession
    ListeTitres Cmb_Titre
    ListeFratrie Cmb_CodeFratrie
End Sub

Public Sub ListeAssuranceMaladie(Boite As ComboBox)
'Place la liste des régimes d'assurance maladie dans la combo

    With Dat_AssuranceMaladie
        .RecordSource = "SELECT ASMA_DESIG, ASMA_CODE FROM ASMA ORDER BY ASMA_DESIG"
        .Refresh
        Boite.Clear
        While Not Dat_AssuranceMaladie.Recordset.EOF
            Boite.AddItem (.Recordset.Fields("ASMA_DESIG") + " <" + CStr(.Recordset.Fields("ASMA_CODE")) + ">")
            .Recordset.MoveNext
        Wend
    End With
End Sub

Public Sub ListeProfessions(Boite As ComboBox)
'Place la liste des professions dans la combo
    With Dat_Profession
        .RecordSource = "SELECT PROF_CODE, PROF_DESI1, PROF_DESI2 FROM PROF ORDER BY PROF_DESI1, PROF_DESI2"
        .Refresh
        Boite.Clear
        While Not .Recordset.EOF
            If .Recordset.Fields("PROF_DESI1") <> "Null" Then
                If .Recordset.Fields("PROF_DESI2") <> "Null" Then
                    Boite.AddItem (.Recordset.Fields("PROF_DESI1") + " " + .Recordset.Fields("PROF_DESI2") + " <" + CStr(.Recordset.Fields("PROF_CODE")) + ">")
                Else
                    Boite.AddItem (.Recordset.Fields("PROF_DESI1") + " <" + CStr(.Recordset.Fields("PROF_CODE")) + ">")
                End If
            Else
                If .Recordset.Fields("PROF_DESI2") <> "Null" Then
                    Boite.AddItem (.Recordset.Fields("PROF_DESI1") + " <" + CStr(.Recordset.Fields("PROF_CODE")) + ">")
                Else
                    Boite.AddItem ("<" + CStr(.Recordset("PROF_CODE")) + ">")
                End If
            End If
            .Recordset.MoveNext
        Wend
    End With
End Sub



Public Sub ListeTitres(Boite As ComboBox)
'Place la liste des titres dans la combo
        Boite.AddItem ("Mr")
        Boite.AddItem ("Mme")
        Boite.AddItem ("Mlle")
End Sub

Public Sub ListeFratrie(Boite As ComboBox)
'Place la liste des code de Fratrie dans la combo
    Boite.AddItem ("Demandeur <DE>")
    Boite.AddItem ("Conjoint  <CO>")
    Boite.AddItem ("Enfant au foyer <EF>")
    Boite.AddItem ("Enfant hors foyer <EH>")
    Boite.AddItem ("Autre personne <AP>")
End Sub



Private Sub Txt_Adr_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez l'adresse de l'employeur de l'individu"
    End If
End Sub

Private Sub Txt_Adr1Pre_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez l'adresse précédente de l'individu"
    End If
End Sub

Private Sub Txt_Adr2Pre_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez l'adresse précédente de l'individu"
    End If
End Sub

Private Sub Txt_CPEmployeur_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez le code postal de l'employeur de l'individu"
    End If
End Sub

Private Sub Txt_CPpre_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez le code postal précédent de l'individu"
    End If
End Sub

Private Sub Txt_DepartementNaissance_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez le département de naissance de l'individu"
    End If
End Sub

Private Sub Txt_Nom_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez le nom de l'individu"
    End If
End Sub

Private Sub Txt_NomEmployeur_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez le nom de l'employeur de l'individu"
    End If
End Sub

Private Sub Txt_Prenom_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez le prénom de l'individu"
    End If
End Sub

Private Sub Txt_TelEmployeur_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez le numéro de téléphone de l'employeur de l'individu"
    End If
End Sub

Private Sub Txt_VilleEmployeur_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez la ville de l'employeur de l'individu"
    End If
End Sub

Private Sub Txt_VilleNaissance_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez la ville de naissance de l'individu"
    End If
End Sub

Private Sub Txt_VillePre_Gotfocus()
    If (stb_Individu.Panels(4).Text = "CREATION") Or (stb_Individu.Panels(4).Text = "MODIFICATION") Then
        stb_Individu.SimpleText = "Entrez la ville de domiciliation précédente de l'individu"
    End If
End Sub
