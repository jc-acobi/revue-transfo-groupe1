<?php
require_once __DIR__ . '/migrate.php';

/**
 * Connexion à la base de données de l'environnement courant.
 *
 * - Les identifiants viennent d'un fichier serveur (/var/www/config/db.php),
 *   jamais du dépôt.
 * - La base dépend de l'environnement, déduit du chemin :
 *     /var/www/apps/groupeN              -> groupeN_prod
 *     /var/www/workshop/groupeN/dev      -> groupeN_dev
 *     /var/www/workshop/groupeN/binome1  -> groupeN_binome1
 *     /var/www/workshop/groupeN/binome2  -> groupeN_binome2
 *   Elle est créée automatiquement si elle n'existe pas encore.
 * - Les migrations en attente (dossier migrations/) sont appliquées automatiquement.
 *
 * Utilisation dans une page :
 *   require_once __DIR__ . '/db.php';
 *   $bdd = db();
 *   $lignes = $bdd->query("SELECT * FROM ma_table")->fetchAll();
 */
function db(): PDO
{
    static $pdo = null;
    if ($pdo !== null) {
        return $pdo;
    }

    $configPath = '/var/www/config/db.php';
    if (!is_file($configPath)) {
        throw new RuntimeException("Configuration de la base introuvable sur le serveur.");
    }
    $cfg = require $configPath;

    if (!preg_match('/groupe\d+/', __DIR__, $g)) {
        throw new RuntimeException("Impossible de déterminer le groupe.");
    }
    $groupe = $g[0];

    if (preg_match('#[\\\\/]apps[\\\\/]#', __DIR__)) {
        $env = 'prod';
    } elseif (preg_match('#[\\\\/](binome1|binome2|dev)[\\\\/]?$#', __DIR__, $e)) {
        $env = $e[1];
    } else {
        $env = 'dev';
    }
    $schema = $groupe . '_' . $env;

    $pdo = new PDO(
        "mysql:host={$cfg['host']};charset=utf8mb4",
        $cfg['user'],
        $cfg['pass'],
        [
            PDO::ATTR_ERRMODE            => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        ]
    );

    // Crée la base de l'environnement si besoin, puis s'y place.
    $pdo->exec("CREATE DATABASE IF NOT EXISTS `{$schema}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci");
    $pdo->exec("USE `{$schema}`");

    // Applique les migrations en attente sur cette base.
    migrate($pdo, __DIR__ . '/migrations');

    return $pdo;
}
