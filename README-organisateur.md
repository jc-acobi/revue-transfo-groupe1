# Atelier vibe coding — gabarit (guide organisateur)

> Ce dépôt est le **gabarit générique**. Il ne se déploie jamais lui-même. On en
> fabrique un dépôt par groupe (`groupe1` … `groupe6`).
> Ce fichier est destiné à **l'organisateur**, pas aux participants.

## Principe

Chaque **groupe** dispose de son propre dépôt. Un groupe = 4 personnes = 2 binômes.
Chaque binôme travaille de son côté sans gêner l'autre, puis tout le monde se rejoint
sur une intégration commune, et enfin la version finale est publiée sur le portail interne.

## Les 4 environnements (par groupe)

| Espace de travail | Branche | Dossier sur le serveur | Adresse |
|---|---|---|---|
| Binôme 1 | `binome1` | `/var/www/workshop/<groupe>/binome1` | `interne.acobi.fr/workshop/<groupe>/binome1/` |
| Binôme 2 | `binome2` | `/var/www/workshop/<groupe>/binome2` | `interne.acobi.fr/workshop/<groupe>/binome2/` |
| Intégration du groupe | `dev` | `/var/www/workshop/<groupe>/dev` | `interne.acobi.fr/workshop/<groupe>/dev/` |
| **Portail (version finale)** | `main` | `/var/www/apps/<groupe>` | `interne.acobi.fr/apps/<groupe>/` |

Le flux : `binome1` / `binome2` → `dev` → `main`. Seule la branche `main` est visible sur le portail.

## Intégration au portail

Le portail (`interne.acobi.fr`) scanne automatiquement `/var/www/apps/*/manifest.json`.
Dès que la version finale est publiée dans `/var/www/apps/<groupe>` (avec un `manifest.json`
où `visible: true`), une carte apparaît sur le portail. **Aucune modification du portail n'est nécessaire.**

## Créer le dépôt d'un groupe

1. Sur GitHub, marquer ce dépôt comme *template repository* (Settings → « Template repository »).
2. Cliquer **« Use this template »** pour créer `revue-transfo-groupe2` (puis `…3`, `…4`, etc.).
3. Cloner le nouveau dépôt en local, puis l'instancier :

   ```powershell
   ./nouveau-groupe.ps1 -Numero 2 -Pousser
   ```

   Le script remplace les repères `__GROUPE__` / `__GROUPE_LABEL__`, crée les espaces de
   travail `dev`, `binome1`, `binome2`, et publie le tout.

> Les noms de dossiers serveur sont dérivés du numéro de groupe (`groupe2`, `groupe3`…),
> pas du nom du dépôt GitHub. Le garde-fou `if: github.repository != 'jc-acobi/revue-transfo'`
> dans chaque workflow empêche le gabarit de se déployer.

## Prérequis serveur (une fois par groupe, avant l'atelier)

1. Renseigner les secrets du dépôt du groupe : `VPS_HOST`, `VPS_USER`, `VPS_PASSWORD`.
2. Sur le serveur, préparer chaque dossier comme un clone positionné sur la bonne branche :

   ```bash
   git clone -b main    <url-du-depot-groupe> /var/www/apps/groupe2
   git clone -b dev      <url-du-depot-groupe> /var/www/workshop/groupe2/dev
   git clone -b binome1  <url-du-depot-groupe> /var/www/workshop/groupe2/binome1
   git clone -b binome2  <url-du-depot-groupe> /var/www/workshop/groupe2/binome2
   ```

3. Vérifier que le serveur web sert ces dossiers aux adresses indiquées plus haut.

## Repères remplacés à l'instanciation

| Repère | Devient (ex. groupe 2) | Présent dans |
|---|---|---|
| `__GROUPE__` | `groupe2` | `CLAUDE.md`, workflows |
| `__GROUPE_LABEL__` | `Groupe 2` | `manifest.json`, `index.php` |
