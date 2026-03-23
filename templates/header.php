<?php $loggedInUser = current_user(); ?>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title><?php echo htmlspecialchars($pageTitle, ENT_QUOTES, 'UTF-8'); ?></title>
    <link rel="stylesheet" href="assets/styles.css">
</head>
<body>
    <div class="page-shell">
        <header class="site-header">
            <div>
                <p class="eyebrow">Cotizador PHP</p>
                <h1><?php echo htmlspecialchars($pageTitle, ENT_QUOTES, 'UTF-8'); ?></h1>
            </div>
            <?php if ($loggedInUser): ?>
                <div class="user-panel">
                    <span>Sesión: <?php echo htmlspecialchars($loggedInUser['display_name'], ENT_QUOTES, 'UTF-8'); ?></span>
                    <a class="button button-secondary" href="logout.php">Cerrar sesión</a>
                </div>
            <?php endif; ?>
        </header>
        <main class="content-card">
