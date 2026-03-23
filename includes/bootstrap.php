<?php

if (session_status() === PHP_SESSION_NONE) {
    session_start();
}

$users = require __DIR__ . '/../config/users.php';
$catalog = require __DIR__ . '/../config/catalog.php';

require_once __DIR__ . '/../lib/auth.php';
require_once __DIR__ . '/../lib/view.php';
require_once __DIR__ . '/../lib/catalog.php';
require_once __DIR__ . '/../lib/quote.php';
require_once __DIR__ . '/../lib/counter.php';
require_once __DIR__ . '/../lib/docx.php';

if (!isset($_SESSION['quote_items'])) {
    $_SESSION['quote_items'] = [];
}
