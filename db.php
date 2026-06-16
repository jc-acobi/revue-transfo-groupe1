<?php
/**
 * Connexion à la base de données du groupe.
 *
 * Les identifiants ne sont PAS ici : ils sont lus depuis un fichier de
 * configuration présent uniquement sur le serveur (/var/www/config/db.php),
 * jamais versionné. Le nom de la base est déduit automatiquement du dossier
 * (groupe1, groupe2, …), de sorte que chaque groupe parle à sa propre base.
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

    if (!preg_match('/groupe\d+/', __DIR__, $m)) {
        throw new RuntimeException("Impossible de déterminer la base du groupe.");
    }
    $schema = $m[0];

    $pdo = new PDO(
        "mysql:host={$cfg['host']};dbname={$schema};charset=utf8mb4",
        $cfg['user'],
        $cfg['pass'],
        [
            PDO::ATTR_ERRMODE            => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        ]
    );

    return $pdo;
}
