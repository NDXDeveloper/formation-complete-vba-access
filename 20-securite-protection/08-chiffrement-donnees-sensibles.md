🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.8. Chiffrement des données sensibles dans les tables

Le chiffrement par mot de passe vu en [section 20.2](02-protection-mot-de-passe.md) protège le fichier au repos : tant que la base est fermée, son contenu est illisible. Mais dès qu'elle est ouverte, toutes les données redeviennent lisibles en clair pour quiconque utilise l'application ou parvient à atteindre les données ouvertes — via une table liée, une copie de sauvegarde ou le back-end accessible. Pour protéger certaines informations particulièrement sensibles au-delà de ce périmètre, on chiffre sélectivement les données elles-mêmes, champ par champ. Ainsi, même si quelqu'un lit le contenu d'une table, les valeurs concernées restent inintelligibles sans la clé. C'est une couche supplémentaire de défense en profondeur, qui vise typiquement les mots de passe applicatifs, les données personnelles et les coordonnées bancaires.

## La distinction fondamentale : hachage et chiffrement

Avant tout, il faut distinguer deux opérations que l'on confond souvent, mais qui répondent à des besoins opposés.

Le hachage est une transformation à sens unique : on ne peut pas retrouver la valeur d'origine à partir de son empreinte. C'est exactement ce qu'il faut pour les mots de passe, car on n'a jamais besoin de relire un mot de passe : on a seulement besoin de vérifier qu'une saisie correspond. On stocke donc l'empreinte du mot de passe, et pour vérifier on recalcule l'empreinte de la saisie pour la comparer à celle qui est stockée.

Le chiffrement, lui, est réversible : avec la bonne clé, on retrouve la valeur d'origine. C'est ce qu'il faut pour les données que l'on doit pouvoir relire plus tard, comme un numéro de téléphone, un IBAN ou un identifiant à transmettre.

La règle qui en découle est sans ambiguïté : un mot de passe se hache, il ne se chiffre pas ; une donnée que l'on doit récupérer se chiffre, elle ne se hache pas.

## Le vrai problème : la gestion de la clé

Tout chiffrement réversible suppose une clé, et c'est là que réside la difficulté majeure dans une application Access côté client. Où stocker cette clé ? Si elle figure dans le code VBA — même compilé en `.accde` — ou dans une table, elle reste, en dernier ressort, accessible à quelqu'un disposant d'un accès suffisant au fichier. Il n'existe pas, dans Access, de coffre-fort à secrets sûr.

Il faut donc être honnête, dans le prolongement de tout ce chapitre : le chiffrement de champs dans Access élève réellement le niveau de protection contre certaines menaces — lecture occasionnelle, vol d'une sauvegarde, fouille via une table liée — mais il n'est pas inviolable face à un attaquant déterminé capable d'extraire la clé de l'application. Pour une protection véritablement forte de données sensibles, la bonne réponse demeure un serveur conçu pour cela, comme SQL Server avec son chiffrement transparent des données, sa fonctionnalité Always Encrypted ou un chiffrement au niveau colonne (voir [chapitre 23](../23-migration-interoperabilite/README.md)).

## Ce qu'il ne faut surtout pas faire

Quelques erreurs sont à proscrire absolument. Ne jamais stocker un mot de passe en clair, bien sûr. Ne jamais confondre encodage et chiffrement : convertir une valeur en Base64 ne la protège en rien, c'est une simple représentation que n'importe qui peut décoder. Et surtout, ne jamais inventer son propre algorithme de chiffrement ni se contenter d'une obfuscation triviale comme un OU exclusif avec une clé fixe : ces bricolages donnent une illusion de sécurité tout en étant cassés en quelques minutes. La cryptographie est un domaine où l'improvisation est dangereuse.

## Hacher les mots de passe

VBA ne fournit aucune fonction de hachage native. Pour calculer une empreinte robuste, il faut s'appuyer sur l'API cryptographique de Windows (CryptoAPI ou la génération CNG plus récente), appelée par des déclarations d'API (voir [chapitre 22](../22-api-windows-integration-office/README.md)). La bibliothèque CAPICOM autrefois employée a été retirée des versions modernes de Windows et ne doit plus servir.

Le principe d'un stockage correct de mot de passe combine trois éléments. D'abord un sel, c'est-à-dire une valeur aléatoire propre à chaque mot de passe, stockée à côté de l'empreinte ; il empêche que deux utilisateurs ayant le même mot de passe produisent la même empreinte, et déjoue les attaques par tables précalculées. Ensuite une fonction lente et itérée — un dérivateur de clé comme PBKDF2 appliqué de nombreuses fois — plutôt qu'un simple SHA-256 en une passe, afin de rendre les attaques par force brute coûteuses. Enfin la vérification par recalcul et comparaison.

Le code ci-dessous illustre l'intégration de cette logique. Les fonctions cryptographiques `CalculerHashPBKDF2` et `GenererSelAleatoire` n'y sont volontairement pas implémentées : elles doivent envelopper l'API Windows. Ne réimplémentez jamais un algorithme cryptographique vous-même.

```vba
' Les primitives ci-dessous s'appuient sur l'API cryptographique de Windows
' (CNG / CryptoAPI, voir chapitre 22). Elles ne sont PAS à réécrire à la main.

Public Sub DefinirMotDePasse(ByVal utilisateurID As Long, ByVal nouveau As String)
    ' Crée un sel aléatoire, calcule l'empreinte et stocke sel + empreinte.
    Dim sel As String
    sel = GenererSelAleatoire(16)                 ' 16 octets aléatoires
    Dim empreinte As String
    empreinte = CalculerHashPBKDF2(nouveau, sel)  ' fonction lente et salée

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT Sel, Empreinte FROM tblUtilisateurs WHERE UtilisateurID = " _
        & utilisateurID, dbOpenDynaset)
    rs.Edit
    rs!Sel = sel
    rs!Empreinte = empreinte
    rs.Update
    rs.Close
    Set rs = Nothing: Set db = Nothing
End Sub

Public Function VerifierMotDePasse(ByVal saisie As String, _
                                   ByVal selStocke As String, _
                                   ByVal empreinteStockee As String) As Boolean
    ' Recalcule l'empreinte de la saisie avec le sel stocké, puis compare.
    Dim empreinteCalculee As String
    empreinteCalculee = CalculerHashPBKDF2(saisie, selStocke)
    VerifierMotDePasse = (StrComp(empreinteCalculee, empreinteStockee, _
                                  vbBinaryCompare) = 0)
End Function
```

On notera que l'on ne déchiffre jamais un mot de passe : on ne fait que recalculer et comparer. Si un utilisateur oublie son mot de passe, on ne le lui « retrouve » pas, on lui en attribue un nouveau.

## Chiffrer des données récupérables

Pour les données que l'application doit pouvoir relire, on recourt au chiffrement réversible, en gardant à l'esprit le problème de la clé exposé plus haut.

Sur Windows, un mécanisme particulièrement adapté aux secrets propres à un utilisateur est l'API de protection des données, dite DPAPI (`CryptProtectData` et `CryptUnprotectData`). Son grand intérêt est qu'elle dérive elle-même la clé des informations d'identification de l'utilisateur ou de la machine : on n'a donc aucune clé à gérer ni à stocker soi-même. Sa contrepartie est décisive pour une base partagée : par défaut, seules les données chiffrées par un utilisateur donné peuvent être déchiffrées par ce même utilisateur. DPAPI convient donc parfaitement pour stocker un secret personnel — par exemple le mot de passe de connexion propre au poste courant — mais ne permet pas que plusieurs utilisateurs lisent une même donnée chiffrée. Une portée « machine » existe, mais elle lie alors le déchiffrement à l'ordinateur, si bien qu'une copie de la base ne se déchiffre plus ailleurs.

Pour une donnée chiffrée que de multiples utilisateurs doivent pouvoir lire, il faut au contraire une clé partagée, ce qui ramène inévitablement au problème de son stockage et donc à la limite de sécurité déjà soulignée.

Du point de vue de l'intégration dans Access, le schéma est toujours le même : on chiffre au moment de l'écriture, on stocke le texte chiffré dans le champ, et on déchiffre au moment de l'affichage. On emploie commodément un champ texte long pour conserver le chiffré sous forme de texte, associé à un contrôle indépendant pour la saisie et l'affichage en clair.

```vba
' Chiffrer / Dechiffrer s'appuient sur l'API Windows (DPAPI ou CNG, chapitre 22)
' et renvoient / acceptent du texte (par exemple du Base64). Pas d'algorithme maison.

Private Sub Form_Current()
    ' Affiche la valeur déchiffrée dans un contrôle indépendant
    Me.txtIBAN_Clair = Dechiffrer(Nz(Me!IBAN_Chiffre, ""))
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    ' Enregistre la valeur chiffrée à partir du contrôle de saisie
    Me!IBAN_Chiffre = Chiffrer(Nz(Me.txtIBAN_Clair, ""))
End Sub
```

Le chiffrement des champs a toutefois un coût fonctionnel qu'il faut accepter. Un champ chiffré ne peut plus être recherché, trié ni indexé sur sa valeur en clair, et toute recherche partielle devient impossible. Il existe bien un chiffrement dit déterministe, qui produit toujours le même chiffré pour une même valeur et autorise donc une recherche par égalité ; mais il révèle quelles lignes partagent la même valeur, ce qui constitue une fuite d'information. Le chiffrement aléatoire, plus sûr, interdit en revanche toute recherche. Il faut donc réserver le chiffrement aux champs qui n'ont pas besoin d'être interrogés.

## Données personnelles et conformité

Le chiffrement des données personnelles ou sensibles s'inscrit dans les mesures techniques attendues en matière de protection des données, notamment au regard du RGPD. Il participe à limiter les conséquences d'une fuite. Cette préoccupation rejoint l'audit : comme indiqué en [section 20.7](07-audit-acces-modifications.md), veillez à ne jamais consigner ces données sensibles en clair dans les journaux, sous peine de réintroduire par la trace ce que le chiffrement vise à protéger.

## Recommandations

En synthèse, la marche à suivre se résume simplement. Pour un mot de passe : le hacher avec un sel et une fonction lente, ne jamais le chiffrer ni le stocker en clair. Pour un secret propre à un poste ou à un utilisateur : recourir à DPAPI, qui dispense de gérer une clé. Pour une donnée partagée à relire par plusieurs utilisateurs : la chiffrer en assumant la faiblesse de la gestion de clé côté client, et, si l'enjeu est réel, déporter le stockage vers un serveur. Dans tous les cas : s'appuyer sur l'API cryptographique de Windows ou une bibliothèque éprouvée, jamais sur un algorithme improvisé, et ne pas confondre encodage et chiffrement.

## Ce qu'il faut retenir

Le chiffrement des données sensibles ajoute une protection ciblée, distincte du chiffrement global du fichier : il garde certaines valeurs inintelligibles même lorsque la base est ouverte ou copiée. Son efficacité repose sur deux principes — hacher les mots de passe, chiffrer les données récupérables — et bute sur une limite incontournable côté client : la clé doit bien résider quelque part. Dans Access, ce chiffrement élève utilement le niveau de protection sans jamais être inviolable, et n'a de sens qu'intégré aux autres couches du chapitre. Pour des données réellement critiques, la migration vers un serveur reste la réponse de fond.

Ainsi s'achève ce chapitre consacré à la sécurité. Aucune des mesures étudiées — chiffrement du fichier, verrouillage du code, confiance des macros, gestion des droits, audit, chiffrement des données — ne suffit isolément ; c'est leur combinaison, calibrée selon le risque réel, qui constitue une protection cohérente. Et lorsque les exigences dépassent ce qu'Access peut raisonnablement offrir, l'orientation vers un véritable serveur de bases de données demeure la décision la plus sûre.

---


⏭️ [21. Déploiement et distribution](/21-deploiement-distribution/README.md)
