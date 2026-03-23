<?php

require __DIR__ . '/includes/bootstrap.php';

if (is_logged_in()) {
    header('Location: cotizador.php');
    exit;
}

header('Location: login.php');
exit;
