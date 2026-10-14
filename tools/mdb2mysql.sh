#!/usr/bin/env bash
# Transfère toutes les tables d'un fichier Access (.mdb/.accdb) vers une base MySQL/MariaDB.
#
# Prérequis : mdbtools (mdb-tables, mdb-schema, mdb-export) et le client mysql.
#   Debian/Ubuntu : apt install mdbtools mariadb-client      macOS : brew install mdbtools mysql-client
#
# Usage :
#   tools/mdb2mysql.sh <fichier.mdb> <base_mysql> [--drop]
#
# Connexion MySQL via variables d'environnement (valeurs par défaut entre parenthèses) :
#   MYSQL_HOST (127.0.0.1)  MYSQL_PORT (3306)  MYSQL_USER (root)  MYSQL_PWD (vide)
#
# La base cible est créée si besoin (utf8mb4). Les tables sont recopiées telles quelles, avec leurs noms
# d'origine : utilisez une base dédiée (ex. adoptions_legacy), pas la base de la nouvelle application.
# Les requêtes enregistrées d'Access (vues) et les états ne sont pas transférés, seulement les tables.
# --drop : supprime d'abord les tables existantes du même nom (DROP TABLE IF EXISTS).
set -euo pipefail

if [[ $# -lt 2 ]]; then
    sed -n '2,17p' "$0" | sed 's/^# \{0,1\}//'
    exit 1
fi

MDB=$1
DB=$2
DROP=${3:-}

[[ -f $MDB ]] || { echo "Fichier introuvable : $MDB" >&2; exit 1; }
[[ $DB =~ ^[A-Za-z0-9_]+$ ]] || { echo "Nom de base invalide : $DB (lettres, chiffres et _ uniquement)" >&2; exit 1; }
for cmd in mdb-tables mdb-schema mdb-export mysql; do
    command -v "$cmd" >/dev/null || { echo "Commande manquante : $cmd" >&2; exit 1; }
done

MYSQL=(mysql --host="${MYSQL_HOST:-127.0.0.1}" --port="${MYSQL_PORT:-3306}" --user="${MYSQL_USER:-root}" --default-character-set=utf8mb4 --binary-mode)
[[ -n ${MYSQL_PWD:-} ]] && export MYSQL_PWD

"${MYSQL[@]}" -e "CREATE DATABASE IF NOT EXISTS \`$DB\` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
MYSQL+=("$DB")

# Les relations d'Access (clés étrangères) ne sont pas recréées : MySQL en refuse certaines (clé référencée non unique)
# et elles sont inutiles pour une copie de travail. Les tables, colonnes et index sont conservés.
{
    echo "SET FOREIGN_KEY_CHECKS=0; SET NAMES utf8mb4; SET sql_mode='NO_AUTO_VALUE_ON_ZERO';"
    if [[ $DROP == --drop ]]; then
        mdb-tables -1 "$MDB" | while IFS= read -r t; do echo "DROP TABLE IF EXISTS \`$t\`;"; done
    fi
    mdb-schema "$MDB" mysql | grep -viE '^(DROP TABLE|ALTER TABLE .* ADD CONSTRAINT .* FOREIGN KEY)'
} | "${MYSQL[@]}"

# Données, table par table (dates au format MySQL, transaction par table pour la vitesse).
total=0
while IFS= read -r table; do
    n=$(mdb-export "$MDB" "$table" | tail -n +2 | wc -l | tr -d ' ')
    printf '%-40s %8s ligne(s)\n' "$table" "$n"
    {
        echo "SET FOREIGN_KEY_CHECKS=0; SET NAMES utf8mb4; SET autocommit=0;"
        mdb-export -I mysql -D '%Y-%m-%d %H:%M:%S' -b strip "$MDB" "$table"
        echo "COMMIT;"
    } | "${MYSQL[@]}"
    total=$((total + n))
done < <(mdb-tables -1 "$MDB")

echo "Terminé : $total ligne(s) transférée(s) vers la base $DB."
