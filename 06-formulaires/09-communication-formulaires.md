🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.9. Communication entre formulaires (sans couplage fort)

Dans une application Access, les formulaires ont constamment besoin d'échanger des informations : une liste ouvre une fiche de détail, une fiche notifie une liste qu'un enregistrement a changé, un assistant transmet ses choix à l'écran suivant. La manière naïve de le faire — un formulaire qui lit et écrit directement dans les contrôles d'un autre — fonctionne, mais crée un **couplage fort** : des dépendances fragiles, difficiles à maintenir et à faire évoluer.

Cette section présente les techniques permettant à des formulaires de communiquer **proprement**, en réduisant ce couplage : propriétés et méthodes publiques, événements personnalisés, variables de session.

---

## 6.9.1. Le problème du couplage fort

Le couplage fort survient lorsqu'un formulaire référence directement les **contrôles internes** d'un autre par leur nom :

```vba
' Anti-patron : F_Detail écrit directement dans un contrôle de F_Liste
Forms!F_Liste!txtTotal.Value = Me.Montant.Value
```

Cette ligne paraît simple, mais elle accumule les défauts :

- **F_Liste doit être ouvert** : sinon l'instruction déclenche une erreur.
- **Le nom `txtTotal` est figé** : renommer ce contrôle casse silencieusement F_Detail.
- **Les dépendances sont invisibles** : rien dans F_Liste n'indique qu'un autre formulaire écrit dans ses entrailles.
- **La réutilisation devient impossible** : F_Detail ne fonctionne plus sans F_Liste, et inversement.

L'objectif du découplage est de faire communiquer les formulaires à travers une **interface stable** (des propriétés, des méthodes, des événements) plutôt qu'à travers leurs contrôles internes.

---

## 6.9.2. Communication directe via la collection Forms (et ses limites)

L'accès direct par la collection `Forms`, présenté au chapitre 6.2, reste la voie la plus immédiate :

```vba
Forms!F_Liste!txtTotal.Value = 100      ' notation point d'exclamation
Forms("F_Liste").Controls("txtTotal") = 100   ' notation explicite
```

Cette approche est **acceptable** dans les cas simples : petites applications, ou relation étroite et durable entre deux formulaires (notamment parent/sous-formulaire, chapitre 6.4). Elle devient problématique dès que l'application grandit ou que les formulaires doivent être réutilisés indépendamment.

Pour toute communication non triviale, les mécanismes décrits ci-après sont préférables.

---

## 6.9.3. Propriétés publiques : exposer une interface propre

Le module de classe d'un formulaire peut exposer des **propriétés publiques** (`Property Let`, `Get`, `Set` — chapitre 16.2). Le formulaire appelant manipule alors ces propriétés sans rien savoir des contrôles internes : c'est le formulaire récepteur qui décide, en interne, comment traduire la valeur reçue.

Dans le module de classe du formulaire récepteur :

```vba
' --- Module de classe de F_Detail ---
Private m_lngClientID As Long

Public Property Let ClientID(ByVal lngValeur As Long)
    m_lngClientID = lngValeur
    ChargerEnregistrement m_lngClientID   ' réaction interne au changement
End Property

Public Property Get ClientID() As Long
    ClientID = m_lngClientID
End Property
```

Le formulaire appelant ne dialogue qu'avec la propriété `ClientID` :

```vba
DoCmd.OpenForm "F_Detail"
Forms!F_Detail.ClientID = Me.lstClients.Value   ' interface stable
```

Renommer un contrôle interne de F_Detail, changer sa logique de chargement : rien de cela n'affecte l'appelant, qui ne connaît que la propriété. C'est le principe « programmer vers une interface, pas vers une implémentation ».

> À l'ouverture d'un formulaire, le passage de paramètres via `OpenArgs` (chapitre 6.6) remplit un rôle voisin, mais limité à une chaîne et au seul moment de l'ouverture. Les propriétés publiques fonctionnent à tout moment et acceptent n'importe quel type.

---

## 6.9.4. Méthodes publiques : déclencher une action

Sur le même principe, un formulaire peut exposer des **procédures publiques** que d'autres formulaires invoquent pour déclencher une action, sans toucher à ses contrôles.

```vba
' --- Dans F_Liste : méthode publique de rafraîchissement ---
Public Sub RafraichirListe()
    Me.lstClients.Requery
End Sub
```

Après avoir enregistré une modification, F_Detail demande simplement à F_Liste de se rafraîchir :

```vba
Private Sub cmdEnregistrer_Click()
    EnregistrerEnregistrement

    If CurrentProject.AllForms("F_Liste").IsLoaded Then
        Forms!F_Liste.RafraichirListe
    End If
End Sub
```

Le test `IsLoaded` (chapitre 4.4) évite l'erreur si F_Liste n'est pas ouvert. Cette approche est plus propre que de requêter directement la liste depuis l'extérieur (`Forms!F_Liste!lstClients.Requery`), car elle laisse F_Liste maître de sa propre logique. Elle conserve néanmoins une dépendance : F_Detail nomme explicitement F_Liste.

---

## 6.9.5. Événements personnalisés (WithEvents) : le découplage maximal

Pour qu'un formulaire **notifie** sans connaître son destinataire, on recourt aux **événements personnalisés** (`RaiseEvent` / `WithEvents`), traités en profondeur au chapitre 16.4. Le principe inverse la dépendance : l'émetteur déclare et déclenche un événement ; ce sont les abonnés qui décident d'y réagir. L'émetteur ignore totalement qui l'écoute.

L'émetteur (F_Detail) déclare un événement et le déclenche :

```vba
' --- Module de classe de F_Detail ---
Public Event EnregistrementModifie(ByVal lngID As Long)

Private Sub cmdEnregistrer_Click()
    EnregistrerEnregistrement
    RaiseEvent EnregistrementModifie(m_lngClientID)
End Sub
```

L'abonné (F_Liste) détient une référence à F_Detail déclarée `WithEvents`, et traite l'événement :

```vba
' --- Module de classe de F_Liste ---
Private WithEvents m_frmDetail As Form_F_Detail

Private Sub cmdOuvrirDetail_Click()
    Set m_frmDetail = New Form_F_Detail   ' instancie le formulaire de détail
    m_frmDetail.Visible = True
End Sub

Private Sub m_frmDetail_EnregistrementModifie(ByVal lngID As Long)
    Me.lstClients.Requery                 ' réagit à la notification
End Sub
```

Ici, **F_Detail ne mentionne jamais F_Liste**. On pourrait ajouter d'autres abonnés sans modifier une ligne de F_Detail : c'est le patron Observateur (chapitre 16.7), idéal pour les notifications « un émetteur, plusieurs destinataires ». En contrepartie, sa mise en œuvre est plus avancée que les techniques précédentes.

---

## 6.9.6. TempVars : un état partagé sans référence directe

Les **`TempVars`** (chapitre 15.7) sont des variables globales de session, accessibles depuis n'importe quel objet. Elles offrent un « tableau noir » partagé où un formulaire dépose une information qu'un autre relit, sans qu'aucun ne référence l'autre.

```vba
' Formulaire A dépose une valeur
TempVars!ClientCourant = Me.lstClients.Value

' Formulaire B la relit, plus tard, sans connaître A
Me.txtClient.Value = TempVars!ClientCourant
```

Ce mécanisme convient particulièrement au **contexte applicatif global** : utilisateur connecté, paramètres généraux, sélection courante partagée par plusieurs écrans. Il est à employer avec mesure : un usage excessif disperse l'état de l'application dans des variables globales difficiles à tracer.

---

## 6.9.7. Le cas particulier parent / sous-formulaire

La relation entre un formulaire et son sous-formulaire est, par nature, étroite ; elle est détaillée au chapitre 6.4. Le sous-formulaire accède à son parent via `Me.Parent`, et le parent à son sous-formulaire via le contrôle conteneur puis `.Form`.

Même dans ce contexte, le découplage reste bénéfique : plutôt qu'un sous-formulaire qui écrit directement dans un contrôle du parent (`Me.Parent!txtTotal = …`), on privilégie l'appel d'une **méthode publique** exposée par le parent (`Me.Parent.MettreAJourTotal …`), qui garde la maîtrise de son propre affichage.

---

## 6.9.8. Choisir le bon mécanisme

| Besoin | Mécanisme | Chapitre |
|---|---|---|
| Transmettre un contexte à l'ouverture (parent → enfant) | `OpenArgs` | 6.6 |
| Affecter des valeurs à un formulaire ouvert, sans connaître ses contrôles | Propriétés publiques | 16.2 |
| Déclencher une action sur un autre formulaire | Méthode publique | — |
| Notifier sans connaître le destinataire | Événements personnalisés (`WithEvents`) | 16.4 |
| Partager un état applicatif global | `TempVars` | 15.7 |
| Relation parent / sous-formulaire | `Me.Parent` + méthode publique | 6.4 |

En règle générale : commencer par la solution la plus simple qui découple suffisamment (propriétés et méthodes publiques), et réserver les **événements personnalisés** aux cas où l'émetteur ne doit véritablement rien savoir de ses destinataires.

---

## 6.9.9. Pièges et bonnes pratiques

- **Programmer vers une interface, pas vers des noms de contrôles** : exposer des propriétés et méthodes publiques plutôt que de laisser l'extérieur manipuler les contrôles internes.
- **Vérifier qu'un formulaire est ouvert** avant de le solliciter : `CurrentProject.AllForms("…").IsLoaded` (chapitre 4.4) prévient les erreurs.
- **Ne pas exposer directement les contrôles** dans les propriétés publiques : encapsuler la logique, afin que l'implémentation interne reste libre d'évoluer.
- **Avec `WithEvents`, ne pas oublier `Set`** : tant que la référence n'est pas affectée (`Set m_frmDetail = …`), aucun événement n'est reçu. Surveiller aussi la durée de vie de l'instance et les références circulaires.
- **Limiter les `TempVars`** au contexte réellement global : un usage abusif transforme l'application en un fouillis d'état partagé difficile à suivre.
- **Centraliser plutôt que disperser** : éviter les références `Forms!X!controle` éparpillées dans tout le code ; les regrouper derrière une interface claire facilite la maintenance.

---

## 6.9.10. Récapitulatif

- Le **couplage fort** — référencer directement les contrôles internes d'un autre formulaire — produit des dépendances fragiles ; le découplage repose sur des **interfaces stables**.
- L'accès direct via la collection `Forms` (chapitre 6.2) reste acceptable pour les cas simples ou la relation parent/sous-formulaire.
- Les **propriétés publiques** (chapitre 16.2) permettent d'affecter des valeurs à un formulaire sans connaître ses contrôles ; les **méthodes publiques** déclenchent des actions en lui laissant la maîtrise de sa logique.
- Les **événements personnalisés** (`RaiseEvent` / `WithEvents`, chapitre 16.4) offrent le découplage maximal : l'émetteur ignore ses destinataires (patron Observateur, chapitre 16.7).
- Les **`TempVars`** (chapitre 15.7) partagent un état global sans référence directe, à réserver au contexte applicatif.
- Toujours vérifier l'ouverture d'un formulaire (`IsLoaded`, chapitre 4.4) avant de le solliciter, et privilégier systématiquement la communication par interface plutôt que par noms de contrôles.

⏭️ [7. États (Reports)](/07-etats-reports/README.md)
