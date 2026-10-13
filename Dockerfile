# syntax=docker/dockerfile:1

# --- Dépendances PHP (sans les outils de développement)
FROM composer:2 AS vendor
WORKDIR /app
COPY composer.json composer.lock symfony.lock ./
RUN composer install --no-dev --no-scripts --no-autoloader --prefer-dist --no-interaction

# --- Image d'exécution : Apache + PHP 8.3
FROM php:8.3-apache AS app
RUN apt-get update \
    && apt-get install -y --no-install-recommends libicu-dev libzip-dev unzip default-mysql-client \
    && docker-php-ext-install -j"$(nproc)" intl pdo_mysql zip opcache \
    && a2enmod rewrite \
    && rm -rf /var/lib/apt/lists/*

ENV APP_ENV=prod \
    APACHE_DOCUMENT_ROOT=/var/www/html/public
RUN sed -ri 's!/var/www/html!${APACHE_DOCUMENT_ROOT}!g' /etc/apache2/sites-available/*.conf /etc/apache2/apache2.conf /etc/apache2/conf-available/*.conf \
    && printf 'opcache.validate_timestamps=0\nopcache.memory_consumption=128\n' > /usr/local/etc/php/conf.d/opcache.ini

COPY --from=composer:2 /usr/bin/composer /usr/bin/composer
WORKDIR /var/www/html
COPY --from=vendor /app/vendor ./vendor
COPY . .
RUN composer dump-autoload --classmap-authoritative --no-dev --no-scripts \
    && mkdir -p var/cache var/log \
    && APP_SECRET=build DATABASE_URL="sqlite:///:memory:" APP_ADMIN_PASSWORD_HASH=x php bin/console cache:warmup --env=prod \
    && chown -R www-data:www-data var

COPY docker/entrypoint.sh /usr/local/bin/entrypoint.sh
RUN chmod +x /usr/local/bin/entrypoint.sh
ENTRYPOINT ["entrypoint.sh"]
CMD ["apache2-foreground"]
EXPOSE 80
