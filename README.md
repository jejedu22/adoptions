# Adoptions

Application de gestion des dossiers d'adoption (agréments, demandes, commissions, recours, interventions
des psychologues et travailleurs sociaux, placements et surveillances), réécrite sous **Symfony 7.4 LTS**
(PHP ≥ 8.2, Doctrine ORM, Twig).

L'ancienne application Visual Basic 6 / Access / Crystal Reports est conservée à titre de référence dans `legacy/`.

## Démarrage

```bash
composer install
php bin/console doctrine:migrations:migrate --no-interaction
php bin/console app:init            # paramètres de l'échéancier et types de recours par défaut
php -S 127.0.0.1:8000 -t public     # ou symfony serve
```

Accès protégé par HTTP Basic : utilisateur `admin`, mot de passe initial `changeme`.
Pour le changer : `php -r 'echo password_hash("nouveau", PASSWORD_BCRYPT);'` puis renseigner
`APP_ADMIN_PASSWORD_HASH` dans `.env.local` (entourer la valeur de quotes simples).
La base est SQLite par défaut (`DATABASE_URL` dans `.env`). Pour MySQL, changez cette variable puis créez le schéma avec
`php bin/console doctrine:schema:create` (les fichiers de `migrations/` sont générés pour SQLite).

## Interface

Navigation latérale (Accueil, Dossiers, À traiter, Placements, Commissions, Référentiels), recherche globale (Ctrl K),
thème clair/sombre automatique. Parcours principaux :

- **Accueil** : indicateurs et actions prioritaires (rapports en retard, décisions à saisir, agréments à réactualiser).
- **À traiter** (`/a-traiter`) : toutes les échéances, filtrables ; « Reçu aujourd'hui » enregistre la réception d'un rapport.
- **Dossiers** : filtres par statut, prochaine étape de chaque dossier ; fiche avec frise d'avancement (demande → dossier reçu →
  rapports → commission → agrément), échéancier et onglets ; assistant de création en 4 étapes.
- Le statut et la prochaine étape sont déduits des données (`App\Service\StatutCalculator`), les échéances par `App\Service\TacheFinder`.
  Une maquette statique est conservée dans `docs/maquette-ihm.html`.

## Correspondance avec l'ancienne application

| VB6 | Symfony |
|---|---|
| Menu « Définir » / « Préparer » (listes + fiches) | `/gestion/{slug}` — CRUD générique piloté par `App\Service\Catalogue` |
| Gestion des dossiers / demandes | `/dossiers`, `/demandes/{id}` |
| Décision de la commission, recours | entité `Decision` (provenance : commission, recours administratif, recours contentieux) |
| `ActualiserDonnéesAgrément` (`Adoptions.bas`) | `App\Service\AgrementCalculator` (testé) |
| Échéancier et paramètres | `App\Service\EcheancierCalculator`, entité `Parametre` |
| Interventions, enfants souhaités, placements, surveillances | entités filles de `Demande` / `Placement` |

## Docker (derrière Traefik)

```bash
cp .env.docker.example .env.docker      # renseigner APP_HOST, APP_SECRET, DB_PASSWORD, APP_ADMIN_PASSWORD_HASH
docker compose --env-file .env.docker up -d --build
```

Aucun port n'est publié : l'application est joignable uniquement via Traefik, par les labels du service `app`
(routeur `Host(APP_HOST)`, point d'entrée `websecure`, TLS avec le résolveur `TRAEFIK_CERTRESOLVER`, port interne 80).
Le service `app` rejoint le réseau externe `TRAEFIK_NETWORK` (celui de Traefik, voir `docker network ls`) ; MySQL reste
sur le réseau interne, avec un volume `db_data`. Adaptez les noms d'entrypoint et de résolveur à votre configuration Traefik ;
la redirection HTTP → HTTPS se règle dans Traefik.

Symfony fait confiance aux en-têtes `X-Forwarded-*` (variable `TRUSTED_PROXIES`, par défaut l'adresse du proxy qui se connecte),
pour générer des URL en https et lire la bonne adresse IP cliente.

Au démarrage, `docker/entrypoint.sh` attend la base, crée ou met à jour le schéma
(`doctrine:schema:update --force`) et exécute `app:init`. Le mot de passe par défaut de `admin` est `changeme` : à changer.
Dans un fichier `.env`, les `$` du hash bcrypt doivent être doublés (`$$`).

Import des données Access depuis le conteneur :

```bash
docker compose --env-file .env.docker exec app php bin/console app:import-legacy --source=mysql://user:mdp@hote/adoptions_legacy --check
```

## Import des données de l'ancienne base

Le fichier `legacy/Access/Adoptions.mdb` n'est qu'un frontal : ses tables sont *liées* à un fichier de données
sur le serveur (`\\S100022\adoptions_p\7.0.0\Access\Adoptions.mdb`). C'est ce dernier qu'il faut copier.

```bash
# 1. Copier le fichier de données Access dans une base MySQL dédiée
tools/mdb2mysql.sh Adoptions.mdb adoptions_legacy
# 2. Vérifier les tables et colonnes attendues (rien n'est écrit)
php bin/console app:import-legacy --source=mysql://user:mdp@127.0.0.1/adoptions_legacy --check
# 3. Essai à blanc (tout est annulé), puis import réel
php bin/console app:import-legacy --source=mysql://user:mdp@127.0.0.1/adoptions_legacy --dry-run
php bin/console app:import-legacy --source=mysql://user:mdp@127.0.0.1/adoptions_legacy
```

La correspondance tables → entités est dans `src/Import/LegacyMapping.php`. Elle a été établie à partir du code VB6
(requêtes SQL, formulaires) sans accès aux données : les colonnes absentes de la source sont signalées par `--check` et
ignorées (`--strict` pour échouer à la place). L'import est rejouable (table `legacy_map`) et recalcule les agréments.

## Reste à reprendre depuis `legacy/`
- États Crystal Reports (`legacy/Crystal/*.rpt`) et génération des documents Word.
- Enfants accordés (table `COAC`) et documents Word (`WDOC`) : non importés.
- Contrôles de cohérence détaillés de la saisie des décisions (`booSaisieIncohérenteDécision`) et recalcul complet de l'échéancier.

## Tests

```bash
vendor/bin/phpunit
```
