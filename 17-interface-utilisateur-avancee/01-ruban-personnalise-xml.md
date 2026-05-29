🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.1. Ruban personnalisé (Ribbon) avec XML et callbacks VBA

Le ruban est la première chose que voit l'utilisateur d'une application Access, et par défaut il affiche l'intégralité des commandes d'Access — Création, Données externes, Outils de base de données — dont la quasi-totalité n'a aucun sens pour l'utilisateur final d'une application métier. Remplacer ce ruban générique par un ruban dédié, ne présentant que les actions pertinentes et déclenchant votre propre code, est l'une des étapes les plus visibles de la transformation d'une base Access en logiciel professionnel.

La personnalisation du ruban repose sur deux composants indissociables. D'une part, un document **XML** rédigé dans un langage de balisage appelé **RibbonX**, qui décrit la structure visuelle du ruban : ses onglets, ses groupes, ses boutons. D'autre part, des **fonctions de rappel** (*callbacks*) écrites en VBA, vers lesquelles le XML pointe pour exécuter une action lorsque l'utilisateur clique, ou pour interroger l'état d'un contrôle. Le XML décrit *ce qui s'affiche* ; le VBA décrit *ce qui se passe*. Cette section détaille les deux faces, ainsi que le mécanisme propre à Access qui les relie.

## Le mécanisme propre à Access : la table USysRibbons

C'est ici que la personnalisation du ruban en Access diffère fondamentalement de celle d'Excel ou de Word. Dans les autres applications Office, le XML du ruban est embarqué dans le fichier du document (qui est une archive Open XML). Access, lui, ne fonctionne pas ainsi : il lit les définitions de ruban dans une **table système nommée `USysRibbons`**, qu'il faut créer soi-même car elle n'existe pas par défaut.

Cette table comporte au minimum deux colonnes :

- `RibbonName` — un texte court servant à nommer et identifier le ruban (par exemple `RubanPrincipal`).
- `RibbonXML` — un champ Mémo (texte long) contenant le balisage RibbonX complet.

Le préfixe `USys` est un préfixe réservé aux objets système : la table sera donc masquée par défaut dans le volet de navigation. Pour la voir après l'avoir créée, il faut activer l'affichage des objets système (clic droit sur le titre du volet de navigation → *Options de navigation* → cocher *Afficher les objets système*). Quel que soit son affichage, dès lors qu'une table porte exactement ce nom et contient ces colonnes, Access lit son contenu **au démarrage** et rend les rubans nommés disponibles dans toute l'application.

Ce point de mécanique a une conséquence pratique à retenir absolument : **toute modification du XML stocké dans `USysRibbons` ne prend effet qu'après fermeture et réouverture de la base de données**, car le ruban est mis en cache au lancement. Modifier le XML et s'attendre à voir le changement immédiatement est l'erreur de débutant la plus courante sur ce sujet.

## Préalable indispensable : activer l'affichage des erreurs de ruban

Avant même d'écrire la première ligne de XML, il faut activer une option sans laquelle le développement du ruban devient un cauchemar. Par défaut, si le XML est mal formé ou si un callback référencé est introuvable, Access **échoue silencieusement** : le ruban personnalisé ne s'affiche tout simplement pas, sans le moindre message d'explication.

Pour obtenir des messages d'erreur explicites, activez l'option *Afficher les erreurs d'interface utilisateur des compléments* : **Fichier → Options → Paramètres du client → Général → cocher « Afficher les erreurs d'interface utilisateur des compléments »**. Une fois cette case cochée, toute anomalie dans le XML (balise non fermée, attribut inconnu, identifiant en double, callback manquant) provoquera l'affichage d'un message indiquant précisément la nature et l'emplacement du problème. C'est l'outil de débogage numéro un du ruban : on l'active dès le début et on le laisse actif pendant tout le développement.

## Anatomie du XML RibbonX

Le balisage RibbonX suit une hiérarchie stricte. La racine est l'élément `<customUI>`, qui déclare un espace de noms (*namespace*) et porte les callbacks globaux. À l'intérieur, l'élément `<ribbon>` contient `<tabs>`, qui contient un ou plusieurs `<tab>` (onglets), chacun contenant des `<group>` (groupes), chacun contenant les contrôles (`<button>`, etc.).

Le choix de l'espace de noms détermine les fonctionnalités disponibles. Deux versions coexistent :

- `http://schemas.microsoft.com/office/2006/01/customui` — la version originale (introduite avec Office 2007), largement compatible et suffisante pour la grande majorité des besoins.
- `http://schemas.microsoft.com/office/2009/07/customui` — une version étendue (Office 2010 et ultérieurs) qui ajoute la personnalisation de la vue Backstage et l'insertion de contrôles dans les onglets intégrés.

Sauf besoin spécifique de ces ajouts, l'espace de noms 2006 est le choix par défaut recommandé pour sa compatibilité maximale. Voici un premier ruban complet et fonctionnel qui servira de fil conducteur :

```xml
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"
          onLoad="RubanCharge">
  <ribbon startFromScratch="false">
    <tabs>
      <tab id="tabGestion" label="Gestion">
        <group id="grpClients" label="Clients">
          <button id="btnNouveauClient"
                  label="Nouveau client"
                  size="large"
                  imageMso="RecordsAddRecord"
                  onAction="OuvrirNouveauClient"/>
          <button id="btnListeClients"
                  label="Liste des clients"
                  size="large"
                  imageMso="ViewsDatasheetView"
                  onAction="OuvrirListeClients"/>
        </group>
        <group id="grpOutils" label="Outils">
          <button id="btnActualiser"
                  label="Actualiser"
                  imageMso="Refresh"
                  onAction="ActualiserDonnees"/>
          <button id="btnSupprimer"
                  label="Supprimer"
                  imageMso="DeleteRows"
                  getEnabled="EstSuppressionPossible"
                  onAction="SupprimerEnregistrement"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

Chaque élément qui porte un identifiant utilise l'attribut `id`, dont la valeur doit être **unique dans tout le document** (un identifiant dupliqué fait échouer le chargement). Les attributs `label` définissent les textes affichés, `imageMso` l'icône (voir plus loin), et les attributs commençant par `on` ou `get` désignent les callbacks VBA. L'attribut `startFromScratch` sur l'élément `<ribbon>`, ici à `false`, conserve les onglets intégrés d'Access ; nous verrons plus loin l'effet de `true`.

## Installer et activer le ruban

Une fois le XML rédigé, trois opérations restent à accomplir : créer la table `USysRibbons`, y insérer le XML, puis indiquer à Access d'utiliser ce ruban.

La table peut être créée manuellement dans l'interface d'Access (création d'une table avec un champ `RibbonName` en Texte court et un champ `RibbonXML` en Texte long, puis saisie d'une ligne). Mais puisque cette formation porte sur VBA, la voie programmatique est plus intéressante et, surtout, réutilisable au déploiement : une procédure d'installation peut créer la table et y injecter le XML, ce qui permet d'automatiser la mise en place du ruban chez chaque utilisateur. Cette routine s'appuie sur le DDL (chapitre 11) et les objets DAO (chapitre 9) :

```vba
' ============================================
' Module standard : modRubanInstallation
' Crée la table USysRibbons et y enregistre le ruban.
' À lancer une fois ; nécessite ensuite une fermeture/
' réouverture de la base pour que le ruban soit pris en compte.
' ============================================
Option Compare Database
Option Explicit

Private Const NOM_RUBAN As String = "RubanPrincipal"

Public Sub InstallerRuban()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb

    ' 1. Créer la table système si elle n'existe pas encore.
    If Not TableExiste("USysRibbons") Then
        db.Execute "CREATE TABLE USysRibbons " & _
                   "(ID AUTOINCREMENT PRIMARY KEY, " & _
                   " RibbonName TEXT(255), " & _
                   " RibbonXML MEMO)", dbFailOnError
    End If

    ' 2. Purger une éventuelle définition portant le même nom.
    db.Execute "DELETE FROM USysRibbons WHERE RibbonName = '" & _
               NOM_RUBAN & "'", dbFailOnError

    ' 3. Insérer le XML (Recordset utilisé plutôt qu'un INSERT SQL :
    '    le champ Mémo peut être volumineux et contenir des apostrophes).
    Set rs = db.OpenRecordset("USysRibbons", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!RibbonName = NOM_RUBAN
    rs!RibbonXML = XmlDuRuban()
    rs.Update
    rs.Close

    Set rs = Nothing
    Set db = Nothing

    MsgBox "Ruban installé." & vbCrLf & _
           "Fermez puis rouvrez la base, puis sélectionnez '" & _
           NOM_RUBAN & "' dans les options de la base courante.", _
           vbInformation, "Installation du ruban"
End Sub

' Renvoie le XML du ruban sous forme de chaîne.
Private Function XmlDuRuban() As String
    Dim x As String
    x = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    x = x & "<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" onLoad=""RubanCharge"">" & vbCrLf
    x = x & "  <ribbon startFromScratch=""false"">" & vbCrLf
    x = x & "    <tabs>" & vbCrLf
    x = x & "      <tab id=""tabGestion"" label=""Gestion"">" & vbCrLf
    x = x & "        <group id=""grpClients"" label=""Clients"">" & vbCrLf
    x = x & "          <button id=""btnNouveauClient"" label=""Nouveau client"" size=""large"" imageMso=""RecordsAddRecord"" onAction=""OuvrirNouveauClient""/>" & vbCrLf
    x = x & "        </group>" & vbCrLf
    x = x & "      </tab>" & vbCrLf
    x = x & "    </tabs>" & vbCrLf
    x = x & "  </ribbon>" & vbCrLf
    x = x & "</customUI>"
    XmlDuRuban = x
End Function

' Indique si une table existe dans la base courante.
Private Function TableExiste(ByVal nomTable As String) As Boolean
    On Error Resume Next
    TableExiste = (Len(CurrentDb.TableDefs(nomTable).Name) > 0)
End Function
```

Concaténer le XML ligne à ligne dans `XmlDuRuban` est fastidieux pour un ruban volumineux ; une alternative consiste à coller le XML directement dans la table via l'interface, ou à le charger depuis un fichier externe avec le `FileSystemObject` (section 22.10). L'approche programmatique reste néanmoins la plus pratique pour une installation automatisée.

Une fois le ruban présent dans `USysRibbons` et la base rouverte, il faut encore l'**activer**, à l'un de deux niveaux :

Au niveau de **l'application entière** : **Fichier → Options → Base de données active → champ « Nom du ruban »**, dans lequel une liste déroulante propose les noms présents dans `USysRibbons`. Le ruban choisi devient le ruban par défaut de l'application.

Au niveau d'un **formulaire ou état particulier** : chaque formulaire et état possède une propriété **« Nom du ruban »** (onglet *Autres* de la feuille de propriétés). Lorsqu'on y indique un nom de ruban, ce ruban s'affiche dès que l'objet prend le focus, ce qui permet de présenter des commandes différentes selon le contexte — un ruban d'édition pour un formulaire de saisie, un ruban d'impression pour un état.

## Le mécanisme des callbacks : relier le XML au VBA

C'est le cœur du dispositif. Les attributs du XML qui commencent par `on` ou `get` ne contiennent pas de code : ils contiennent le **nom d'une procédure VBA** qu'Access appellera au moment voulu. Ces procédures doivent respecter trois règles impératives : être **publiques**, résider dans un **module standard** (jamais dans un module de classe ni de formulaire), et porter une **signature exacte** correspondant au type d'événement attendu. Un callback dont la signature ne correspond pas (mauvais nombre ou type de paramètres) sera considéré comme introuvable et fera échouer le contrôle.

La signature varie selon le rôle du callback. Le tableau suivant récapitule les plus courantes :

| Attribut XML | Rôle | Signature VBA |
|--------------|------|---------------|
| `onLoad` (sur `customUI`) | Capturer l'objet ruban au chargement | `Sub Nom(ribbon As IRibbonUI)` |
| `onAction` (button) | Réagir à un clic | `Sub Nom(control As IRibbonControl)` |
| `onAction` (toggleButton, checkBox) | Réagir à un basculement | `Sub Nom(control As IRibbonControl, pressed As Boolean)` |
| `getEnabled` | Activer/désactiver dynamiquement | `Sub Nom(control As IRibbonControl, ByRef enabled)` |
| `getVisible` | Afficher/masquer dynamiquement | `Sub Nom(control As IRibbonControl, ByRef visible)` |
| `getLabel` | Définir le texte dynamiquement | `Sub Nom(control As IRibbonControl, ByRef label)` |
| `getImage` | Fournir une image dynamiquement | `Sub Nom(control As IRibbonControl, ByRef image)` |
| `onChange` (editBox, comboBox) | Réagir à une saisie | `Sub Nom(control As IRibbonControl, text As String)` |
| `getItemCount` (comboBox, dropDown) | Nombre d'éléments | `Sub Nom(control As IRibbonControl, ByRef count)` |
| `getItemLabel` (comboBox, dropDown) | Libellé d'un élément | `Sub Nom(control As IRibbonControl, index As Integer, ByRef label)` |

Le paramètre `control As IRibbonControl` est systématiquement présent (sauf pour `onLoad`) : il identifie le contrôle à l'origine de l'appel et expose notamment sa propriété `.Id`, utile lorsqu'un même callback sert plusieurs contrôles. Les callbacks `get*` reçoivent un dernier paramètre **par référence** (`ByRef`) que l'on renseigne pour communiquer la réponse à Access (la valeur d'activation, le texte, l'image…).

Les types `IRibbonUI` et `IRibbonControl` proviennent de la bibliothèque **Microsoft Office Object Library**, normalement déjà référencée dans tout projet Access (section 2.5). Si cette référence venait à manquer, ou pour un code volontairement portable, on peut déclarer ces paramètres en `As Object` (liaison tardive, section 2.6), au prix de la perte de l'autocomplétion.

Voici le module de callbacks correspondant au ruban d'exemple :

```vba
' ============================================
' Module standard : modRubanCallbacks
' Procédures appelées par le ruban (toutes Public).
' ============================================
Option Compare Database
Option Explicit

' Référence vers l'objet ruban, conservée au niveau du module
' pour pouvoir le réactualiser dynamiquement par la suite.
Private mRuban As IRibbonUI

' --- onLoad : Access transmet l'objet ruban une seule fois,
'     au chargement. On le mémorise pour un usage ultérieur. ---
Public Sub RubanCharge(ribbon As IRibbonUI)
    Set mRuban = ribbon
End Sub

' --- onAction des boutons : ouverture de formulaires via DoCmd (ch. 5) ---
Public Sub OuvrirNouveauClient(control As IRibbonControl)
    On Error GoTo Erreur
    DoCmd.OpenForm "frmClient", , , , acFormAdd
    Exit Sub
Erreur:
    MsgBox "Impossible d'ouvrir le formulaire : " & Err.Description, _
           vbExclamation
End Sub

Public Sub OuvrirListeClients(control As IRibbonControl)
    On Error GoTo Erreur
    DoCmd.OpenForm "frmListeClients"
    Exit Sub
Erreur:
    MsgBox "Impossible d'ouvrir la liste : " & Err.Description, _
           vbExclamation
End Sub

Public Sub ActualiserDonnees(control As IRibbonControl)
    On Error Resume Next
    If Not Screen.ActiveForm Is Nothing Then
        Screen.ActiveForm.Requery
    End If
End Sub

Public Sub SupprimerEnregistrement(control As IRibbonControl)
    On Error Resume Next
    If Not Screen.ActiveForm Is Nothing Then
        If MsgBox("Supprimer l'enregistrement courant ?", _
                  vbYesNo + vbQuestion) = vbYes Then
            DoCmd.RunCommand acCmdDeleteRecord
        End If
    End If
End Sub

' --- getEnabled : interrogé par Access pour savoir si le bouton
'     « Supprimer » doit être actif. On renseigne le paramètre ByRef. ---
Public Sub EstSuppressionPossible(control As IRibbonControl, ByRef enabled)
    On Error Resume Next
    enabled = Not (Screen.ActiveForm Is Nothing)
End Sub
```

On remarque que les callbacks `onAction` sont, dans l'immense majorité des cas, de simples passerelles vers l'objet **DoCmd** (chapitre 5) : ouvrir un formulaire, lancer une commande, exécuter une requête. Le ruban est l'interface ; la logique reste dans le code applicatif.

## Les contrôles courants

RibbonX propose une palette de contrôles couvrant la plupart des besoins. Le `<button>` déclenche une action sur clic. Le `<toggleButton>` et la `<checkBox>` représentent un état activé/désactivé et transmettent ce dernier via le paramètre `pressed` de leur `onAction`. L'`<editBox>` offre une zone de saisie de texte. La `<comboBox>` et la `<dropDown>` proposent une liste de choix, alimentée par les callbacks `getItemCount` et `getItemLabel`. Le `<menu>` regroupe plusieurs commandes dans une liste déroulante, la `<gallery>` présente un choix visuel sous forme de vignettes, le `<separator>` insère une séparation verticale entre contrôles, et le `<labelControl>` affiche un simple texte. Les groupes peuvent être empilés et les boutons dimensionnés via l'attribut `size` (`normal` ou `large`).

Voici un groupe illustrant un `toggleButton` et une `dropDown`, avec leurs callbacks :

```xml
<group id="grpAffichage" label="Affichage">
  <toggleButton id="tglInactifs"
                label="Inclure les inactifs"
                imageMso="FilterToggleFilter"
                onAction="BasculerInactifs"/>
  <dropDown id="ddTri"
            label="Trier par"
            getItemCount="TriNombre"
            getItemLabel="TriLibelle"
            onAction="TriChoisi"/>
</group>
```

```vba
' toggleButton : le paramètre 'pressed' donne le nouvel état.
Public Sub BasculerInactifs(control As IRibbonControl, pressed As Boolean)
    ' On stocke le choix dans une variable de session (TempVars, sec. 15.7)
    ' que le reste de l'application pourra consulter.
    TempVars!AfficherInactifs = pressed
    If Not Screen.ActiveForm Is Nothing Then Screen.ActiveForm.Requery
End Sub

' dropDown : Access demande d'abord le nombre d'éléments...
Public Sub TriNombre(control As IRibbonControl, ByRef count)
    count = 3
End Sub

' ...puis le libellé de chacun (appelé une fois par index).
Public Sub TriLibelle(control As IRibbonControl, index As Integer, ByRef label)
    Select Case index
        Case 0: label = "Nom"
        Case 1: label = "Date d'inscription"
        Case 2: label = "Statut"
    End Select
End Sub

' Et réagit au choix de l'utilisateur.
Public Sub TriChoisi(control As IRibbonControl, _
                     selectedId As String, selectedIndex As Integer)
    ' selectedIndex vaut 0, 1 ou 2 selon l'élément retenu.
    MsgBox "Tri demandé : élément n°" & selectedIndex
End Sub
```

## Les images des contrôles

Trois approches existent pour illustrer un contrôle. La plus simple, et de loin la plus utilisée, consiste à réutiliser les **icônes intégrées d'Office** via l'attribut `imageMso`, en indiquant le nom interne de l'icône (`Refresh`, `FileSave`, `DeleteRows`, `RecordsAddRecord`…). Des milliers d'icônes sont ainsi disponibles sans aucun fichier à fournir ; leurs noms se découvrent via les galeries publiées en ligne ou en exportant une personnalisation du ruban réalisée dans l'interface d'Access. Aucune liste n'est reproduite ici, mais ces icônes couvrent l'essentiel des besoins courants.

Pour des **images personnalisées**, deux callbacks entrent en jeu. L'attribut `image="monImage"` associé au callback global `loadImage` (déclaré sur `<customUI>`) fait appeler ce dernier une fois pour chaque identifiant d'image rencontré, à charge pour lui de renvoyer un objet `IPictureDisp`. L'attribut `getImage`, défini sur un contrôle précis, permet quant à lui de fournir une image par contrôle et de la rafraîchir dynamiquement (par invalidation, voir ci-dessous) :

```vba
' loadImage : appelé une fois par identifiant d'image du XML.
Public Sub ChargerImage(imageID As String, ByRef image)
    Set image = ImageDepuisFichier(CurrentProject.Path & "\img\" & imageID & ".bmp")
End Sub
```

La fabrication d'un `IPictureDisp` à partir d'un fichier dépasse le cadre de cette section : `LoadPicture` (de la bibliothèque `stdole`) convient pour les formats BMP et ICO, mais le PNG avec transparence exige une fonction utilitaire s'appuyant sur GDI+. Un tel chargeur d'image réutilisable figure parmi les modèles de code de l'annexe K. Pour la plupart des applications, les icônes `imageMso` intégrées suffisent et évitent entièrement cette complexité.

## Rubans dynamiques : modifier le ruban en cours d'exécution

Jusqu'ici, le ruban est figé après son chargement. Pour le faire évoluer pendant l'exécution — griser un bouton selon le contexte, changer un libellé, masquer un groupe — il faut conjuguer deux mécanismes : les callbacks `get*` et l'**invalidation**.

Le principe est le suivant. Les callbacks `getEnabled`, `getVisible`, `getLabel`, `getImage` ne sont pas appelés en permanence : Access les interroge **une seule fois**, au moment de l'affichage du contrôle, puis met le résultat en cache. Pour forcer Access à les réinterroger, on **invalide** le contrôle (ou le ruban entier), ce qui le marque comme « périmé » et déclenche un nouvel appel aux callbacks correspondants.

L'invalidation s'effectue via l'objet `IRibbonUI`, capturé une fois pour toutes dans le callback `onLoad` et mémorisé au niveau du module (`mRuban` dans nos exemples). Cet objet expose deux méthodes essentielles : `Invalidate` (réinterroge tous les contrôles) et `InvalidateControl "id"` (réinterroge un contrôle précis, plus efficace).

Le scénario typique : lorsqu'un événement de l'application change le contexte (par exemple le passage d'un formulaire à un autre, ou le déplacement sur un enregistrement), on invalide les contrôles concernés pour que leur état se mette à jour :

```vba
' À appeler depuis l'application quand le contexte évolue
' (par ex. dans l'événement Current d'un formulaire, sec. 8.8) :
Public Sub RafraichirRuban()
    If Not mRuban Is Nothing Then
        ' Réinterroge uniquement le bouton « Supprimer » :
        ' son callback getEnabled (EstSuppressionPossible) sera rappelé.
        mRuban.InvalidateControl "btnSupprimer"
    End If
End Sub
```

La conservation de l'objet `mRuban` est ce qui rend tout cela possible : sans la capture opérée dans `onLoad`, on n'aurait aucun moyen de demander à Access de rafraîchir quoi que ce soit. C'est la raison pour laquelle pratiquement tout ruban un tant soit peu dynamique déclare un callback `onLoad`.

## Masquer entièrement le ruban natif

L'attribut `startFromScratch="true"` sur l'élément `<ribbon>` supprime la totalité des onglets intégrés d'Access, ne laissant que votre ruban personnalisé (et la vue Backstage). C'est l'option à retenir pour une application verrouillée, distribuée en environnement de production : l'utilisateur n'a plus accès à aucune commande de conception, et l'application se présente comme un logiciel entièrement maîtrisé.

```xml
<ribbon startFromScratch="true">
  <!-- seuls vos propres onglets subsistent -->
</ribbon>
```

Cette option se combine naturellement avec les autres techniques de verrouillage de l'interface — masquage du volet de navigation, désactivation des touches spéciales, options de démarrage — traitées en section 17.7, et trouve tout son sens lors d'un déploiement via Access Runtime (chapitre 21), où l'environnement de conception est de toute façon absent.

À l'inverse, pour seulement *compléter* les onglets existants plutôt que de repartir d'une page blanche, on peut référencer un onglet ou un groupe intégré par son identifiant interne via l'attribut `idMso`, et y ajouter ses propres contrôles. Cette possibilité, plus avancée et particulièrement développée avec l'espace de noms 2009, permet une intégration fine au ruban natif lorsqu'elle est souhaitée.

## Gérer les erreurs dans les callbacks

Une attention particulière doit être portée à la robustesse des callbacks. Une erreur non interceptée dans un callback `onAction` peut, selon les versions et les paramètres, soit afficher un message brut peu compréhensible, soit échouer silencieusement. Pire, une erreur dans un callback `get*` peut laisser un contrôle dans un état indéterminé. La règle est donc d'**encadrer systématiquement les callbacks par une gestion d'erreurs** (chapitre 13), comme dans les exemples `OuvrirNouveauClient` ou `SupprimerEnregistrement` ci-dessus. Pour les callbacks `get*`, qui ne doivent jamais bloquer l'affichage, un simple `On Error Resume Next` assorti d'une valeur par défaut sûre est généralement préférable à un déroutement vers un gestionnaire.

## Pièges courants

Le piège le plus fréquent, déjà signalé, est d'**oublier de fermer et rouvrir la base** après modification du XML : le ruban étant mis en cache au démarrage, les changements restent invisibles tant que la base n'a pas été relancée. Le deuxième est de **développer sans avoir activé l'affichage des erreurs d'interface** : on se retrouve alors face à un ruban absent, sans aucune indication sur la cause.

Côté callbacks, les erreurs classiques tiennent à la **signature** : un `onAction` de bouton déclaré avec un paramètre `pressed` (qui n'appartient qu'au `toggleButton`), un callback `get*` ayant oublié le paramètre `ByRef` final, ou une procédure placée par mégarde dans un module de classe ou de formulaire plutôt que dans un module standard. Dans tous ces cas, Access considère le callback comme introuvable. Enfin, les **identifiants en double** (`id`) dans le XML font échouer l'ensemble du chargement, d'où l'intérêt d'une convention de nommage rigoureuse (préfixes `tab`, `grp`, `btn`, `tgl`…) pour éviter les collisions.

Un dernier point de vigilance concerne la **dépendance aux callbacks** : si le ruban référence une procédure absente du projet (callback supprimé ou renommé), le contrôle correspondant sera inactif. Toute opération de refactoring (chapitre 24.5) doit donc maintenir la cohérence entre les noms cités dans le XML et les procédures effectivement présentes.

## Synthèse

Personnaliser le ruban en Access repose sur trois piliers : un document **XML RibbonX** décrivant la structure visuelle, des **callbacks VBA** publics, placés dans un module standard et respectant des signatures précises, et la table système **`USysRibbons`** qui, propre à Access, stocke les définitions et les charge au démarrage. Deux réflexes conditionnent la réussite : activer l'option *Afficher les erreurs d'interface utilisateur des compléments* dès le début du développement, et se rappeler que toute modification du XML exige une fermeture/réouverture de la base.

Au-delà du ruban statique, la capture de l'objet `IRibbonUI` dans le callback `onLoad`, combinée aux méthodes `Invalidate` et aux callbacks `get*`, ouvre la voie à des rubans dynamiques dont l'apparence et l'activation suivent le contexte de l'application. Associé à `startFromScratch="true"` et aux autres techniques de verrouillage (section 17.7), le ruban personnalisé devient la pièce maîtresse de l'interface d'une application Access destinée à la production et au déploiement (chapitre 21).

---

> 🔗 **Pour aller plus loin :** les callbacks `onAction` s'appuient massivement sur l'objet DoCmd (chapitre 5) ; la déclaration des types `IRibbonUI`/`IRibbonControl` relève des références COM (section 2.5) et de la liaison précoce/tardive (section 2.6) ; la création de la table `USysRibbons` mobilise le DDL (section 11.4) et DAO (chapitre 9) ; le verrouillage complet de l'interface est traité en section 17.7 ; et un chargeur d'image `IPictureDisp` réutilisable figure à l'annexe K.

⏭️ [17.2. Menus contextuels personnalisés (CommandBars)](/17-interface-utilisateur-avancee/02-menus-contextuels.md)
