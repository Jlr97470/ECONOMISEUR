VERSION 5.00
Begin VB.Form FrmEconomiseur 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DELTA ECONOMISEUR"
   ClientHeight    =   7560
   ClientLeft      =   3930
   ClientTop       =   3285
   ClientWidth     =   9210
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   9210
   Begin VB.DriveListBox drvEconomiseur 
      Height          =   315
      Left            =   0
      TabIndex        =   13
      Top             =   360
      Width           =   9195
   End
   Begin VB.DirListBox dirEconomiseur 
      Height          =   1215
      Left            =   0
      TabIndex        =   12
      Top             =   675
      Width           =   9195
   End
   Begin VB.CommandButton cmdEconomiseur 
      Caption         =   "CONFIGURATION"
      Height          =   390
      Index           =   1
      Left            =   7290
      TabIndex        =   11
      Top             =   6750
      Width           =   1900
   End
   Begin VB.TextBox txtEconomiseur 
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   4815
      TabIndex        =   10
      Top             =   7155
      Width           =   2460
   End
   Begin VB.CommandButton cmdEconomiseur 
      Caption         =   "QUITTER"
      Height          =   390
      Index           =   2
      Left            =   7290
      TabIndex        =   8
      Top             =   7155
      Width           =   1900
   End
   Begin VB.CommandButton cmdEconomiseur 
      Caption         =   "OK"
      Height          =   390
      Index           =   0
      Left            =   7290
      TabIndex        =   7
      Top             =   6345
      Width           =   1900
   End
   Begin VB.TextBox txtEconomiseur 
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   0
      Left            =   4815
      TabIndex        =   4
      Top             =   6735
      Width           =   2445
   End
   Begin VB.CheckBox cbxEconomiseur 
      BackColor       =   &H0000C000&
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   4815
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   6330
      Width           =   2445
   End
   Begin VB.FileListBox filEconomiseur 
      Height          =   3990
      Left            =   0
      Pattern         =   "*.SCR"
      TabIndex        =   1
      Top             =   1890
      Width           =   9210
   End
   Begin VB.Label lblEconomiseur 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "NOM"
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   4
      Left            =   0
      TabIndex        =   9
      Top             =   7155
      Width           =   4785
   End
   Begin VB.Label lblEconomiseur 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "PARAMETRE DE L'ECRAN DE VEILLE"
      ForeColor       =   &H000000FF&
      Height          =   420
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   5895
      Width           =   9195
   End
   Begin VB.Label lblEconomiseur 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "TEMP D'ACTIVATION"
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   6750
      Width           =   4785
   End
   Begin VB.Label lblEconomiseur 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "ACTIVE"
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   6345
      Width           =   4785
   End
   Begin VB.Label lblEconomiseur 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "ECRAN DE VEILLE"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9225
   End
End
Attribute VB_Name = "frmEconomiseur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    Delta Copyright                                                             (19/04/2001)  ***
'******************************************************************************
'***    FORM:                                                                                              ***
'***      Pour selectionner l'economiseur d'ecran et modifier ces parametres ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Afficher l'economiser d'ecran selectionner dans le panneau de          ***
'***      Controle et ces differents parametres                                              ***
'******************************************************************************
'***    PROGRAMMEUR:                                                                              ***
'***      Royer Jean-Laurent                                                                         ***
'******************************************************************************

'******************************************************************************
'***    MODIF :                                                                                            ***
'***      Version 1.0 : 19/04/2001                                                                  ***
'******************************************************************************
Option Explicit

'******************************************************************************
'***    Declaration De Constante Privee                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Qui Defini Les Libelles De La feuille en erreur                    ***
'******************************************************************************
Private Const mconFEUILLENOM = "frmDeltaEconomiseur"                                ' Je me rapelle le nom de la Feuille
Private Const mconFEUILLETYPE = FEUILLEFORM

'******************************************************************************
'***    Evenement                                                                                       ***
'******************************************************************************

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'***      Form_Load()                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour charger les donnes en relation avec l'economiseur                   ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
Private Sub Form_Load()
   ' En cas d'erreur je gere l'erreur
   On Error GoTo Form_Load_Erreur
   ' Je defini le nom de l'application
   Me.Caption = App.ProductName & " V " & App.Major & "." & App.Minor & "." & App.Revision & " Copyright " + App.LegalCopyright
   ' Je lit les donnees de l'economiseur en cours
   LectureEconomiseur
   ' Fin
Form_Load_Exit:
   ' Je sort de la procedure
   Exit Sub
   ' Fin
Form_Load_Erreur:
     ' Je l'ecrit dans le journal
     gfloLogEconomiseur.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, INSTRUCTIONEVENEMENT, "Form_Load", vbNullString, Err
    ' Je continue
    Resume Form_Load_Exit
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'***      drvEconomiseur_Change()                                                                ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour changer le disque courant                                                        ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
Private Sub drvEconomiseur_Change()
   ' En cas d'erreur je gere l'erreur
   On Error GoTo drvEconomiseur_Change_Erreur
   ' Je met a jour la liste des repertoires
   dirEconomiseur.Path = drvEconomiseur.List(drvEconomiseur.ListIndex)
   ' Fin
drvEconomiseur_Change_Exit:
   ' Je sort de la procedure
   Exit Sub
   ' Fin
drvEconomiseur_Change_Erreur:
     ' Je l'ecrit dans le journal
     gfloLogEconomiseur.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEEVENEMENT, "drvEconomiseur_Change", vbNullString, Err
    ' Je continue
    Resume drvEconomiseur_Change_Exit
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'***      dirEconomiseur_Change()                                                                ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour changer le repertoire courant                                                   ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
Private Sub dirEconomiseur_Change()
   ' En cas d'erreur je gere l'erreur
   On Error GoTo dirEconomiseur_Change_Erreur
   ' Je met a jour la liste des repertoires
   filEconomiseur.Path = dirEconomiseur.List(dirEconomiseur.ListIndex)
   ' Fin
dirEconomiseur_Change_Exit:
   ' Je sort de la procedure
   Exit Sub
   ' Fin
dirEconomiseur_Change_Erreur:
     ' Je l'ecrit dans le journal
     gfloLogEconomiseur.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEEVENEMENT, "dirEconomiseur_Change", vbNullString, Err
    ' Je continue
    Resume dirEconomiseur_Change_Exit
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'***      cmdEconomiseur_Click()                                                                 ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour valider les modifs lier a l'economiseur                                     ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Index - Le numero du bouton appuyer                                               ***
'***    SORTIE:                                                                                           ***
'***      Neant                                                                                             ***
'******************************************************************************
Private Sub cmdEconomiseur_Click(Index As Integer)
     ' En cas d'erreur je gere l'erreur
     On Error GoTo cmdEconomiseur_Click_Erreur
     ' Je regarde le numero du bouton appuyer
     Select Case Index
         Case ECONOMISEUROK
               ' C'est le bouton ok pour mise a jour
               ' Je sauvegarde les donnees de l'economiseur selectionner
               Select Case EcritureEconomiseur
                  Case True
                     ' J'ai bien ecrit les info
                     ' Je previent l'user
                     MsgBox "Mise A jour Reussi"
                  Case False
                     ' Je n'ai pas ecrit les info
                     ' Je previent l'user
                     MsgBox "Mise A jour Incomplete"
               End Select
         Case ECONOMISEURCONFIGURATION
               ' C'est le bouton ok pour mise a jour
               ' J'ouvre l'economiseur en mode configuration
               AfficheEconomiseurConfiguration
         Case ECONOMISEURQUITTER
               ' C'est le bouton quitter pour mise a jour
               ' Je ferme la feuille
               Unload Me
   End Select
    ' Fin
cmdEconomiseur_Click_Exit:
    ' Je sort de la procedure
    Exit Sub
    ' Fin
cmdEconomiseur_Click_Erreur:
    ' Je l'ecrit dans le journal
    gfloLogEconomiseur.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEEVENEMENT, "cmdEconomiseur_Click", vbNullString, Err
    ' Je continue
    Resume cmdEconomiseur_Click_Exit
    ' Fin
End Sub

