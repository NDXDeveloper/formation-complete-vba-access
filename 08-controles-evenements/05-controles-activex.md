🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.5. Contrôles ActiveX dans Access

Les **contrôles ActiveX** sont des composants logiciels externes (issus du monde COM/OCX) que l'on peut intégrer à un formulaire pour obtenir des fonctionnalités absentes des contrôles natifs : arborescences, listes enrichies, barres de progression, sélecteurs de fichiers, navigateurs web, etc. Historiquement précieux, ils sont aujourd'hui à **manier avec prudence** : la plupart de leurs usages sont couverts par des contrôles natifs, et ils s'accompagnent de contraintes de déploiement, de compatibilité 32/64 bits et de sécurité non négligeables.

Cette section explique ce qu'est un contrôle ActiveX, comment le programmer, et surtout **pourquoi privilégier les alternatives natives** dans la grande majorité des cas.

---

## 8.5.1. Qu'est-ce qu'un contrôle ActiveX ?

Un contrôle ActiveX est un **composant réutilisable** développé en dehors d'Access, distribué sous forme de fichier `.ocx` ou `.dll`, et **enregistré** dans Windows pour être utilisable par les applications hôtes. Une fois inséré dans un formulaire, il y apporte son interface et son comportement propres.

Access lui-même repose sur cette technologie pour certains de ses contrôles, mais un contrôle ActiveX **tiers** introduit une dépendance externe qu'il faut gérer (enregistrement, version, licence).

---

## 8.5.2. Les contraintes majeures (à connaître avant tout)

Avant d'envisager un contrôle ActiveX, il faut mesurer ses contraintes — elles motivent à elles seules la préférence pour les solutions natives :

- **Enregistrement sur chaque poste** : le composant doit être **enregistré** (via `regsvr32`) sur **toutes** les machines clientes, ce qui exige souvent des droits administrateur. Un composant non enregistré rend le formulaire inutilisable.
- **Compatibilité 32 / 64 bits** : de nombreux ActiveX n'existent qu'en **32 bits** et ne fonctionnent pas avec une version 64 bits d'Access — un problème de compatibilité majeur (chapitre 21.7).
- **Sécurité** : ActiveX est un vecteur d'attaque classique ; certains composants sont bloqués par le centre de gestion de la confidentialité (chapitre 20.5) ou indisponibles selon la configuration.
- **Licence de distribution** : certains contrôles requièrent une licence pour être déployés.
- **Obsolescence** : plusieurs contrôles historiques ont été **retirés** des versions récentes d'Office.

> En pratique : **ne recourir à un ActiveX que lorsqu'aucune solution native n'existe**, et après avoir validé son fonctionnement sur l'environnement cible exact (bitness, version d'Office, paramètres de sécurité).

---

## 8.5.3. Contrôles ActiveX historiquement utilisés

| Contrôle | Usage | Statut actuel |
|---|---|---|
| Calendar Control (`MSCAL.OCX`) | Sélecteur de date | **Retiré** depuis Office 2010 ; remplacé par le sélecteur natif |
| TreeView / ListView (`MSCOMCTL.OCX`) | Arborescences, listes enrichies | Encore utilisé, mais sujet aux problèmes d'enregistrement |
| ProgressBar (`MSCOMCTL.OCX`) | Barre de progression | Remplaçable par une barre native (chapitre 17.5) |
| Common Dialog (`COMDLG32.OCX`) | Boîtes Ouvrir/Enregistrer | À remplacer par l'objet `FileDialog` natif |
| Web Browser Control | Affichage de pages web | Désormais **natif** dans Access |

Le composant `MSCOMCTL.OCX` (Windows Common Controls) mérite une mention : il a connu plusieurs mises à jour de sécurité qui ont parfois **cassé** la compatibilité des applications l'utilisant (« bibliothèque d'objets non enregistrée » après une mise à jour d'Office).

---

## 8.5.4. Ajouter un contrôle ActiveX

L'insertion se fait en mode Création, via la commande **Insérer un contrôle ActiveX**, qui propose la liste des composants **enregistrés** sur la machine de développement. Le composant doit donc y être présent — et le sera-t-il aussi sur les postes cibles (section 8.5.2).

---

## 8.5.5. Programmer un contrôle ActiveX : la propriété Object

Sur le formulaire, le contrôle ActiveX est encapsulé dans un objet `Control` ordinaire, qui expose les propriétés communes (`Name`, `Visible`, `Left`, `Top`…). Mais l'**API spécifique** du composant — ses propriétés, méthodes et collections propres — s'atteint via la propriété **`Object`** :

```vba
' Accéder à l'objet ActiveX sous-jacent d'un TreeView
Me.axTreeView.Object.Nodes.Add , , "racine", "Racine"
Me.axTreeView.Object.Nodes.Add "racine", 4, "enfant", "Élément"
```

La propriété `Object` est donc la clé pour piloter un contrôle ActiveX par code : elle renvoie le composant réel, avec toute sa richesse, là où le `Control` enveloppe n'expose que le strict commun.

---

## 8.5.6. Liaison précoce ou tardive

Comme pour toute bibliothèque COM (chapitre 2.5), on peut référencer la **bibliothèque de types** du contrôle ActiveX pour bénéficier d'une **liaison précoce** : complétion automatique, vérification des types, constantes nommées. À défaut, on travaille en **liaison tardive** via `Object`, sans assistance de l'éditeur (chapitre 2.6).

La liaison précoce facilite le développement, mais impose que la référence soit résolue sur chaque poste — une contrainte de plus à ajouter à celles de la section 8.5.2.

---

## 8.5.7. Gérer les événements d'un contrôle ActiveX

Un contrôle ActiveX expose ses propres **événements**, qui apparaissent dans les listes déroulantes de l'éditeur VBA (module du formulaire). Les procédures correspondantes sont nommées d'après le contrôle et l'événement :

```vba
Private Sub axTreeView_NodeClick(ByVal Node As Object)
    Debug.Print "Nœud cliqué : " & Node.Text
End Sub
```

Le typage des arguments d'événement dépend de la liaison (précoce ou tardive).

---

## 8.5.8. Les alternatives natives (fortement recommandées)

Pour la plupart des besoins ayant historiquement justifié un ActiveX, Access offre désormais une solution **native**, sans dépendance externe :

| Besoin | Alternative native |
|---|---|
| Sélecteur de date | Sélecteur natif (zone de texte liée à un champ Date/Heure) |
| Barre de progression | Barre construite avec des contrôles natifs (chapitre 17.5) ou `SysCmd` |
| Boîte de dialogue de fichier | Objet **`FileDialog`** : `Application.FileDialog(...)` |
| Affichage web | Contrôle navigateur **natif** |
| Graphiques | Contrôle graphique natif (versions récentes) ou automation Excel (chapitre 22.3) |

Exemple de l'objet `FileDialog`, qui remplace avantageusement le Common Dialog ActiveX :

```vba
Dim fd As Object
Set fd = Application.FileDialog(3)   ' 3 = sélecteur de fichier

If fd.Show = -1 Then                 ' -1 si l'utilisateur valide
    Debug.Print fd.SelectedItems(1)
End If
```

Ces solutions natives évitent l'enregistrement sur les postes, les problèmes de bitness et les blocages de sécurité.

---

## 8.5.9. Pièges et bonnes pratiques

- **Préférer le natif** : n'employer un ActiveX que si aucune alternative native n'existe (sections 8.5.8).
- **Vérifier la compatibilité 32/64 bits** : nombre d'ActiveX sont 32 bits uniquement (chapitre 21.7).
- **Anticiper le déploiement** : le composant doit être enregistré sur **chaque** poste, parfois avec des droits administrateur.
- **Tenir compte de la sécurité** : un ActiveX peut être bloqué par le centre de confidentialité (chapitre 20.5).
- **Vérifier la licence** de distribution des composants tiers.
- **Programmer via la propriété `Object`** pour atteindre l'API complète du composant.
- **Tester sur l'environnement cible exact** (version d'Office, bitness, sécurité) avant tout déploiement.
- **Se méfier des composants retirés** (le Calendar Control, par exemple) dans les versions récentes d'Office.

---

## 8.5.10. Récapitulatif

- Un **contrôle ActiveX** est un composant externe (COM/OCX) intégré à un formulaire pour des fonctionnalités absentes nativement, mais qui doit être **enregistré sur chaque poste**.
- Ses **contraintes majeures** — déploiement, compatibilité **32/64 bits** (chapitre 21.7), sécurité (chapitre 20.5), licence, obsolescence — en font une solution de **dernier recours**.
- On programme un ActiveX via sa propriété **`Object`**, qui expose son API complète ; ses événements apparaissent dans le module du formulaire.
- La **liaison précoce** (référence à la bibliothèque de types, chapitre 2.5) facilite le développement mais ajoute une dépendance ; la **liaison tardive** s'appuie sur `Object` (chapitre 2.6).
- Pour la plupart des besoins, des **alternatives natives** existent : sélecteur de date natif, barre de progression maison (chapitre 17.5), objet **`FileDialog`**, contrôle navigateur natif, graphiques natifs ou automation Excel (chapitre 22.3) — à privilégier systématiquement.

⏭️ [8.6. Propriété ControlSource — contrôles liés et indépendants](/08-controles-evenements/06-controlsource-lie-independant.md)
