<?php
require_once __DIR__ . '/db.php';

$bdd = db();

$bdd->exec("CREATE TABLE IF NOT EXISTS g1_app_data (
    data_key VARCHAR(50) PRIMARY KEY,
    data_value LONGTEXT,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
)");

header('Content-Type: application/json; charset=utf-8');

if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    $stmt = $bdd->prepare("SELECT data_value FROM g1_app_data WHERE data_key = 'db'");
    $stmt->execute();
    $row = $stmt->fetch(PDO::FETCH_ASSOC);
    if ($row) {
        echo $row['data_value'];
    } else {
        echo json_encode(['collaborateurs' => [], 'missions' => [], 'clients' => []]);
    }

} elseif ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $body = file_get_contents('php://input');
    if (!json_decode($body)) {
        http_response_code(400);
        echo json_encode(['error' => 'Données invalides']);
        exit;
    }
    $stmt = $bdd->prepare("INSERT INTO g1_app_data (data_key, data_value) VALUES ('db', ?)
        ON DUPLICATE KEY UPDATE data_value = VALUES(data_value), updated_at = CURRENT_TIMESTAMP");
    $stmt->execute([$body]);
    echo json_encode(['ok' => true]);
}
