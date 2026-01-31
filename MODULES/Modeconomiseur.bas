Attribute VB_Name = "modEconomiseur"
'******************************************************************************
'***    Delta Copyright                                                             (19/04/2001)  ***
'******************************************************************************
'***    MODULE:                                                                                          ***
'***        modEconomiseur                                                                          ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      - Declaration De Varai
'***      - Ecrire Et Lire Les Donnée De L'Economiseur D'Ecran                     ***
'***      - Ouvrir Le Panneau De Configuration De L'Economiseur                   ***
'******************************************************************************
'***    PROGRAMMEUR:                                                                              ***
'***      Royer Jean-Laurent                                                                         ***
'******************************************************************************

'******************************************************************************
'***    MODIF :                                                                                            ***
'***      Version 1.0 : 19/04/2001                                                                  ***
'******************************************************************************
Option Explicit                                                                                               ' Je doit etre sur que mes variables on ete declarer

'******************************************************************************
'***    Declaration De Constante Privee                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Qui Defini Les Libelles De La feuille en erreur                    ***
'******************************************************************************
Private Const mconFEUILLENOM = "modEconomiseur"                                       ' Je me rapelle le nom de la Feuille
Private Const mconFEUILLETYPE = FEUILLEMODULE
'******************************************************************************
'***    Declaration De Constante Public                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Qui Defini Les Boutons de la feuilles                                 ***
'******************************************************************************
Public Enum BOUTONECONOMISEURENUM
    ECONOMISEUROK                                                                                ' Le bouton ok
    ECONOMISEURCONFIGURATION                                                                 ' Le bouton pour ouvrir la feuille de configuration
    ECONOMISEURQUITTER                                                                             ' Le bouton quitter
End Enum

'******************************************************************************
'***    Constante Qui Defini Les Labels de la feuilles                                    ***
'******************************************************************************
Public Enum TXTECONOMISEURENUM
   ECONOMISEURDELAY = 0                                                                            ' Le TextBox du temp avant activation
   ECONOMISEURNOM = 1                                                                                ' Le TextBox du nom de l'ecran de veille
End Enum
'******************************************************************************
'***    Declaration De Variable Priver                                                          ***
'******************************************************************************

'******************************************************************************
'***    Variable Pour La  Gestion D'un Fichier Ini                                          ***
'******************************************************************************
Private mfprFicIni As New clsFicPrf                                                                     ' L'object pour la gestion du fichier profiler

'******************************************************************************
'***    Variable Pour La  Gestion Du Registre Windows                                 ***
'******************************************************************************
Private mrgwRegWin As New RegWin                                                              ' L'object pour la gestion du registre windows

'******************************************************************************
'***    Declaration De Variable Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    Variable Pour La  Gestion D'un Fichier Journal                                   ***
'******************************************************************************
Public gfloLogEconomiseur As New clsFicLog                                           ' L'object pour la gestion d'un fichier journal

'******************************************************************************
'***    Declaration De Function Priver                                                          ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'***        Main()                                                                                           ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***        - Neant                                                                                         ***
'***    SORTIE:                                                                                           ***
'***        - Neant                                                                                         ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***        Main()                                                                                           ***
'******************************************************************************
Private Sub Main()
   ' En Cas D'Erreur Je Gere L'Erreur
   On Error GoTo Main_Erreur
   ' Je creer l'object journal
   Set gfloLogEconomiseur = New clsFicLog
   ' J'ouvre La fenetre
   frmEconomiseur.Show
   ' Fin
Main_Exit:
   ' Je Sort De La Procedure
   Exit Sub
   ' Fin
Main_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogEconomiseur.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, INSTRUCTIONPROCEDURE, "Main", vbNullString, Err
   ' Je Continue
   Resume Main_Exit
   ' Fin
End Sub

'******************************************************************************
'***    Declaration De Function Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      LectureEconomiseur() As Boolean                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Lire les donnees de l'economiseur en cours                              ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      LectureEconomiseur - Renvoie si les donnees on ete lu                    ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =LectureEconomiseur()                                                                   ***
'******************************************************************************
Public Function LectureEconomiseur() As Boolean
   Dim fsoSystemFichier As New FileSystemObject                                          ' L'object de gestion des repertoires et fichiers
   Dim fldRepertoire As Folder                                                                        ' L'object repertoire de l'economiseur
   Dim filFichier As File                                                                                  ' L'object fichier de l'economiseur
   Dim filFichierListe As File                                                                          ' L'object fichier de la liste de fichier
   Dim strRepSystem As String                                                                       ' Le repertoire system de windows
   Dim strRepWindows As String                                                                     ' Le repertoire de windows
   Dim intListeIndex As Integer                                                                        ' L'index dans la liste de fichier
   Dim vntRegValeur As Variant                                                                       ' La valeur de la valeur du registre windows recuperer
   Dim strRegValeur As String                                                                          ' La valeur de la valeur du registre windows recuperer
   Dim strIniValeur As String                                                                            ' La valeur de la valeur du fichier system.ini
   ' En cas d'erreur je gere l'erreur
   On Error GoTo LectureEconomiseur_Erreur
   ' Je renvoie que j'ai bien lu l'enregistrement
   LectureEconomiseur = True
   ' Je recupere le repertoire system
   strRepSystem = fsoSystemFichier.GetSpecialFolder(SystemFolder)
   ' Je recupere le repertoire windows
   strRepWindows = fsoSystemFichier.GetSpecialFolder(WindowsFolder)
   ' Je vais beaucoup utiliser les objets de la form
   With frmEconomiseur
      ' J'applique le filtre des extension des economiseur
      .filEconomiseur.Pattern = "*.SCR"
      ' Je recupere la valeur de la cle qui defini le repertoire d'economiseur et le nom de l'economiseur
      Select Case mrgwRegWin.LectureValeur(HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE", vntRegValeur)
         Case True
            ' c'est une valeur
            ' Je regarde la valeur
            Select Case vntRegValeur
                Case "(Aucun)", vbNullString
                    ' Aucun ecran de veille selectionner
                    ' Je met la valeur
                    strRegValeur = vbNullString
                Case Else
                    ' Un ecran de veille seletionner
                    ' Je met la valeur
                    strRegValeur = vntRegValeur
            End Select
         Case False
            ' c'est une valeur null
            ' Je met la valeur a zero
            strRegValeur = vbNullString
      End Select
      ' Je recupere la valeur de la cle qui defini le repertoire d'economiseur et le nom de l'economiseur
      Select Case mfprFicIni.LectureValeur(strRepWindows & "\" & "SYSTEM.INI", "BOOT", "SCRNSAVE.EXE", strIniValeur, vbNullString)
         Case True
            ' c'est une valeur
            ' Je regarde la valeur
            Select Case strIniValeur
                Case "(Aucun)", vbNullString
                    ' Aucun ecran de veille selectionner
                    ' Je met la valeur
                    strIniValeur = vbNullString
                Case Else
                    ' Un ecran de veille seletionner
                    ' Je met la valeur
                    strIniValeur = strIniValeur
            End Select
         Case False
            ' c'est une valeur null
            ' Je met la valeur a zero
            strIniValeur = vbNullString
      End Select
      ' Je regarde la valeur
      Select Case vbNullString
          Case Is <> strRegValeur
                ' C'est une valeur de base de registre
                ' Je recupere l'object fichier de l'economiseur actuel
                Set filFichier = fsoSystemFichier.GetFile(strRegValeur)
               ' Je defini le disque courant
               .drvEconomiseur.Drive = filFichier.Drive
               ' Je defini le repertoire courant
               .dirEconomiseur.Path = filFichier.ParentFolder.Path
          Case Is <> strIniValeur
                ' C'est une valeur de fichier ini
                ' Je recupere l'object fichier de l'economiseur actuel
                Set filFichier = fsoSystemFichier.GetFile(strIniValeur)
               ' Je defini le disque courant
               .drvEconomiseur.Drive = filFichier.Drive
               ' Je defini le repertoire courant
               .dirEconomiseur.Path = filFichier.ParentFolder.Path
         Case Else
               ' C'est aucunne de ces valeurs
                ' Je recupere l'object repertoire system
                Set fldRepertoire = fsoSystemFichier.GetFolder(strRepSystem)
               ' Je defini le disque courant
               .drvEconomiseur.Drive = fldRepertoire.Drive
               ' Je defini le repertoire courant
               .dirEconomiseur.Path = fldRepertoire.Path
      End Select
      ' Je regarde si j'ai une valeur
      Select Case strRegValeur + strIniValeur
         Case vbNullString
            ' Je n'ai pas de valeur
         Case Else
            ' J'ai une valeur
            ' Je recherche l'economiseur actuel
            For intListeIndex = 0 To .filEconomiseur.ListCount - 1
                ' Je recupere l'object fichier de la liste des fichiers
                Set filFichierListe = fsoSystemFichier.GetFile(.filEconomiseur.Path & "\" & .filEconomiseur.List(intListeIndex))
                 ' Je regarde si c'est le bon economiseur
                 If filFichierListe.ShortName = filFichier.ShortName Then
                       ' C'est le bon fichier
                       ' Je l'affiche a l'user
                       .filEconomiseur.ListIndex = intListeIndex
                       ' Je sort de la boucle
                       Exit For
                 End If
            Next
      End Select
      ' Je recupere la valeur de la cle qui defini si l'economiseur est actif
      Select Case mrgwRegWin.LectureValeur(HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaveActive", vntRegValeur)
         Case True
            ' c'est une valeur
            ' Je met la valeur
            strRegValeur = vntRegValeur
         Case False
            ' c'est une valeur null
            ' Je met la valeur a zero
            strRegValeur = vbNullString
      End Select
      ' Je regarde la valeur
      Select Case strRegValeur
          Case "0"
             ' L'economiseur n'est pas actif
             ' Je l'affiche a l'user
             .cbxEconomiseur.Value = vbUnchecked
          Case "1"
             ' L'economiseur est actif
             ' Je l'affiche a l'user
             .cbxEconomiseur.Value = vbChecked
      End Select
      ' Je recupere la valeur de la cle qui defini le temp avant affichage de l'economiseur
      Select Case mrgwRegWin.LectureValeur(HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaveTimeOut", vntRegValeur)
         Case True
            ' c'est une valeur
            ' Je met la valeur
            strRegValeur = vntRegValeur
         Case False
            ' c'est une valeur null
            ' Je met la valeur a zero
            strRegValeur = vbNullString
      End Select
      ' Je l'affiche a l'user
      .txtEconomiseur(ECONOMISEURDELAY).Text = strRegValeur
   End With
   ' Fin
LectureEconomiseur_Exit:
   ' Je sort de la procedure
   Exit Function
   ' Fin
LectureEconomiseur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogEconomiseur.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, INSTRUCTIONFONCTION, "LectureEconomiseur", vbNullString, Err
   ' Je renvoie que je n'est pas lu l'enregistrement
   LectureEconomiseur = False
   ' Je continue
   Resume LectureEconomiseur_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      EcritureEconomiseur() As Boolean                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour ecrire les donnees de l'economiseur en cours                           ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      EcritureEconomiseur - Renvoie si les donnees on ete lu                   ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =EcritureEconomiseur()                                                                   ***
'******************************************************************************
Public Function EcritureEconomiseur() As Boolean
   Dim fsoSystemFichier As New FileSystemObject                                          ' L'object de gestion des repertoires et fichiers
   Dim filFichierListe As File                                                                           ' L'object fichier de la liste de fichier
   Dim strRepSystem As String                                                                        ' Le repertoire system de windows
   Dim strRepWindows As String                                                                      ' Le repertoire de windows
   Dim vntRegValeur As Variant                                                                       ' La valeur de la valeur du registre windows recuperer
   Dim strRegValeur As String                                                                          ' La valeur de la valeur du registre windows recuperer
   Dim strIniValeur As String                                                                            ' La valeur de la valeur du fichier system.ini
   ' En cas d'erreur je gere l'erreur
   On Error GoTo EcritureEconomiseur_Erreur
   ' Je recupere le repertoire system
   strRepSystem = fsoSystemFichier.GetSpecialFolder(SystemFolder)
   ' Je recupere le repertoire windows
   strRepWindows = fsoSystemFichier.GetSpecialFolder(WindowsFolder)
   ' Je recupere la valeur de la cle qui defini le repertoire d'economiseur et le nom de l'economiseur
   Select Case mrgwRegWin.LectureValeur(HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE", vntRegValeur)
      Case True
         ' c'est une valeur
        ' Je regarde la valeur
        Select Case vntRegValeur
            Case "(Aucun)", vbNullString
                ' Aucun ecran de veille selectionner
                ' Je met la valeur
                strRegValeur = vbNullString
            Case Else
                ' Un ecran de veille seletionner
                ' Je met la valeur
                strRegValeur = vntRegValeur
        End Select
      Case False
         ' c'est une valeur null
         ' Je met la valeur a zero
         strRegValeur = vbNullString
   End Select
   ' Je recupere la valeur de la cle qui defini le repertoire d'economiseur et le nom de l'economiseur
   Select Case mfprFicIni.LectureValeur(strRepWindows & "\" & "SYSTEM.INI", "BOOT", "SCRNSAVE.EXE", strIniValeur, vbNullString)
      Case True
         ' c'est une valeur
            ' Je regarde la valeur
            Select Case strIniValeur
                Case "(Aucun)", vbNullString
                    ' Aucun ecran de veille selectionner
                    ' Je met la valeur
                    strIniValeur = vbNullString
                Case Else
                    ' Un ecran de veille seletionner
            End Select
      Case False
         ' c'est une valeur null
         ' Je met la valeur a zero
         strIniValeur = vbNullString
   End Select
   ' Je vais beaucoup utiliser les objets de la form
   With frmEconomiseur
      ' Je recupere l'object fichier de l'economiseur actuel
      Set filFichierListe = fsoSystemFichier.GetFile(.filEconomiseur.Path & "\" & .filEconomiseur.FileName)
      ' Je regarde la valeur
      Select Case vbNullString
         Case Is <> strRegValeur
            ' C'est une valeur de base de registre
            ' J'ecrit la valeur de la cle qui defini le repertoire d'economiseur et le nom de l'economiseur
            EcritureEconomiseur = mrgwRegWin.EcritureValeur(HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE", CStr(filFichierListe.Path))
         Case Is <> strIniValeur
            ' C'est une valeur de fichier ini
            ' J'ecrit la valeur de la cle qui defini le repertoire d'economiseur et le nom de l'economiseur
            EcritureEconomiseur = mfprFicIni.EcritureValeur("BOOT", "SCRNSAVE.EXE", CStr(filFichierListe.ShortPath), strRepWindows & "\" & "SYSTEM.INI")
         Case Else
            ' C'est une valeur de base de registre
            ' J'ecrit la valeur de la cle qui defini le repertoire d'economiseur et le nom de l'economiseur
            EcritureEconomiseur = mrgwRegWin.EcritureValeur(HKEY_CURRENT_USER, "Control Panel\Desktop", "SCRNSAVE.EXE", CStr(filFichierListe.Path))
      End Select
      ' J'ecrit la valeur de la cle qui defini si l'economiseur est actif
      mrgwRegWin.EcritureValeur HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaveActive", CStr(.cbxEconomiseur.Value)
      ' J'ecrit la valeur de la cle qui defini le temp avant affichage de l'economiseur
      mrgwRegWin.EcritureValeur HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaveTimeOut", CStr(.txtEconomiseur(ECONOMISEURDELAY).Text)
   End With
   ' Fin
EcritureEconomiseur_Exit:
   ' Je sort de la procedure
   Exit Function
   ' Fin
EcritureEconomiseur_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogEconomiseur.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "EcritureEconomiseur", vbNullString, Err
   ' Je renvoie que je n'est pas lu l'enregistrement
   EcritureEconomiseur = False
   ' Je continue
   Resume EcritureEconomiseur_Exit
   ' Fin
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      AfficheEconomiseurConfiguration() As Boolean                                ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'***      Pour Lancer l'economiseur en mode parametre                                 ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***      Neant                                                                                             ***
'***    SORTIE:                                                                                           ***
'***      AfficheEconomiseurConfiguration - Renvoie si les donnees on ete lu ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'***      =AfficheEconomiseurConfiguration()                                                ***
'******************************************************************************
Public Function AfficheEconomiseurConfiguration() As Boolean
   Dim fsoSystemFichier As New FileSystemObject                                           ' L'object de gestion des repertoires et fichiers
   Dim filFichierListe As File                                                                            ' L'object fichier de la liste de fichier
   ' En cas d'erreur je gere l'erreur
   On Error GoTo AfficheEconomiseurConfiguration_Erreur
   ' Je vais beaucoup utiliser les objets de la form
   With frmEconomiseur
      ' Je recupere l'object fichier de l'economiseur actuel
      Set filFichierListe = fsoSystemFichier.GetFile(.filEconomiseur.Path & "\" & .filEconomiseur.FileName)
      ' Je lance une commande shell
      AfficheEconomiseurConfiguration = Shell(filFichierListe.Path & " /C", vbMaximizedFocus)
   End With
   ' Fin
AfficheEconomiseurConfiguration_Exit:
   ' Je sort de la procedure
   Exit Function
   ' Fin
AfficheEconomiseurConfiguration_Erreur:
   ' Je l'ecrit dans le journal
   gfloLogEconomiseur.AjouteErreur App, mconFEUILLETYPE, mconFEUILLENOM, LIBELLEFONCTION, "AfficheEconomiseurConfiguration", vbNullString, Err
   ' Je renvoie que je n'est pas lu l'enregistrement
   AfficheEconomiseurConfiguration = False
   ' Je Continue
   Resume AfficheEconomiseurConfiguration_Exit
   ' Fin
End Function


