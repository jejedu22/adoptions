#!/bin/sh
set -e

# Attend la base de données, crée ou met à jour le schéma, puis charge les données par défaut.
if [ "${SKIP_DB_SETUP:-0}" != "1" ]; then
    echo "Attente de la base de données..."
    i=0
    until php bin/console dbal:run-sql "SELECT 1" >/dev/null 2>&1; do
        i=$((i + 1))
        [ "$i" -gt 60 ] && { echo "Base de données inaccessible." >&2; exit 1; }
        sleep 2
    done
    php bin/console doctrine:schema:update --force --no-interaction
    php bin/console app:init --no-interaction
fi

chown -R www-data:www-data var 2>/dev/null || true
exec docker-php-entrypoint "$@"
