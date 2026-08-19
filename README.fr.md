# Graph App Role Manager

[English version](README.md)

**Version actuelle du dépôt : [v1.2.0](https://github.com/cloudbydefault/graph-app-role-manager/releases/tag/v1.2.0)**  
**Version de l'interface multiplateforme : v1.1.0**

Graph App Role Manager est une interface graphique permettant de consulter, d'attribuer et
de supprimer les permissions d'application Microsoft Graph d'une identité managée ou d'un
Service Principal.

Le dépôt contient deux versions :

- l'interface Python/Tkinter recommandée pour Windows, macOS et les bureaux Linux ;
- une alternative Windows native en PowerShell et Windows Forms.

Les permissions disponibles sont lues directement depuis le Service Principal Microsoft
Graph du tenant. Elles ne sont pas enregistrées dans une liste statique dans le code.

## Nouveautés de la v1.2.0

- Protocole Python/PowerShell persistant fiabilisé avec une sortie UTF-8 déterministe et un décodage défensif côté Python.
- Suppression du préchargement automatique de tous les Service Principals dans l'interface multiplateforme ; la recherche explicite se fait désormais côté serveur.
- Refactorisation de la version Windows native en module PowerShell réutilisable tout en conservant le lanceur `.ps1` existant.
- Ajout du paramètre optionnel `-LogLevel` avec `INFO`, `SUCCESS`, `WARNING` et `ERROR`.
- Isolation de l'état du module natif et initialisation renforcée avec `Set-StrictMode -Version Latest`.
- Extension des tests de régression et du smoke test Windows, avec validation des fichiers `.psm1` dans la CI.

Consultez les [notes de version v1.2.0](https://github.com/cloudbydefault/graph-app-role-manager/releases/tag/v1.2.0) pour le détail.

## Fonctionnalités

- Authentification interactive avec Microsoft Entra ID.
- Recherche des identités managées et Service Principals par nom d'affichage.
- Consultation et filtrage des permissions d'application Microsoft Graph actives.
- Affichage des App Roles Graph déjà attribués à l'identité sélectionnée.
- Attribution de plusieurs permissions en évitant les doublons.
- Suppression des permissions sélectionnées après confirmation explicite.
- Recherche explicite des Service Principals côté serveur dans l'interface multiplateforme, sans précharger toute la liste du tenant.
- Conservation des permissions sélectionnées pendant le filtrage.
- Dialogue dédié au Device Code sous Windows pour éviter les erreurs de Window Handle WAM.
- Interface multiplateforme modernisée avec des contrôles et tableaux plus lisibles.
- Interface Windows native adaptative reprenant la même organisation visuelle.
- Journal d'activité et récapitulatif de chaque opération.
- Aucun mot de passe, secret, certificat ou jeton d'accès enregistré sur le disque.

## Choisir une version

| Version | Usage conseillé | Authentification | Prérequis |
| --- | --- | --- | --- |
| **[Interface Python multiplateforme v1.1.0](cross-platform/graph_app_role_manager.py) — recommandée** | Windows, macOS ou Linux | Device Code sous Windows ; navigateur interactif sous macOS/Linux | Python 3.10+, Tkinter, PowerShell 7, `Microsoft.Graph.Authentication` |
| [Interface PowerShell Windows native](windows/Graph-App-Role-Manager.ps1) | Administrateurs Windows préférant une interface entièrement PowerShell | `Connect-MgGraph` interactif | Windows, PowerShell 7+, modules Microsoft Graph |

La version multiplateforme ne nécessite ni App Registration dédiée, ni Client ID, ni
secret, ni certificat, ni paquet Python supplémentaire.

## Captures d'écran

### Interface multiplateforme (recommandée)

![Interface multiplateforme de Graph App Role Manager sous Windows](docs/images/cross-platform-windows.png)

### Interface PowerShell Windows native

![Interface PowerShell Windows native de Graph App Role Manager](docs/images/native-windows.png)

> **Périmètre de l'interface :** Les interfaces privilégient volontairement la fiabilité
> des opérations Microsoft Graph, la lisibilité, la portabilité et la compatibilité avec
> les composants natifs plutôt qu'une personnalisation visuelle poussée. Tkinter et
> Windows Forms héritent largement de l'apparence du système d'exploitation et offrent
> moins de souplesse visuelle qu'une interface web moderne. Les améliorations purement
> esthétiques ne constituent donc pas l'objectif principal du projet ; la gestion sûre et
> correcte des permissions reste prioritaire.

## Permissions déléguées requises

La session de l'administrateur demande :

- `Application.Read.All`
- `AppRoleAssignment.ReadWrite.All`

Ces permissions sont sensibles. Elles nécessitent un consentement administrateur et un rôle
Microsoft Entra adapté. Consultez
[Permissions et sécurité](docs/permissions-and-safety.md) avant toute utilisation.

## Démarrage rapide

### Version multiplateforme (recommandée)

```bash
pwsh -NoProfile -Command "Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
python cross-platform/graph_app_role_manager.py
```

Conservez `graph_app_role_manager.py`, `graph_backend.ps1` et
`graph-app-role-manager-icon.png` dans le même dossier. Le champ du tenant est facultatif.
Les prérequis par système et le dépannage se trouvent dans le
[guide multiplateforme](docs/cross-platform.md).

### Alternative Windows native

Lanceur standard :

```powershell
pwsh -File .\windows\Graph-App-Role-Manager.ps1
```

Filtrage optionnel du journal :

```powershell
pwsh -File .\windows\Graph-App-Role-Manager.ps1 -LogLevel WARNING
```

Utilisation directe du module :

```powershell
Import-Module .\windows\Graph-App-Role-Manager.psm1
Start-GraphAppRoleManager
```

L'interface native peut proposer l'installation des modules Graph absents pour
l'utilisateur courant, mais demande une confirmation préalable.

## Utilisation prudente

1. Connectez-vous avec un compte administrateur autorisé.
2. Recherchez l'identité cible.
3. Vérifiez son nom, son Object ID, son Application ID et son type.
4. Consultez les permissions existantes.
5. Sélectionnez uniquement les permissions nécessaires au scénario.
6. Relisez la confirmation avant toute attribution ou suppression.
7. Contrôlez le résultat dans l'outil et dans le centre d'administration Microsoft Entra.

## Périmètre

L'outil gère uniquement les **permissions d'application Microsoft Graph** représentées par
des App Role Assignments. Il ne gère pas le consentement délégué des utilisateurs, les rôles
d'annuaire Entra, Azure RBAC ou les autorisations de site SharePoint `Sites.Selected`.

## Documentation

- [Architecture](docs/architecture.md)
- [Installation Windows](docs/windows.md)
- [Installation multiplateforme](docs/cross-platform.md)
- [Permissions et sécurité](docs/permissions-and-safety.md)

Les captures d'écran utilisent un libellé de compte masqué. Évitez de publier de véritables
comptes administrateur, domaines de tenant ou autres informations sensibles du tenant dans
les prochaines captures.

## Licence

Projet distribué sous [licence MIT](LICENSE).
