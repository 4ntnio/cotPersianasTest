<?php

require __DIR__ . '/includes/bootstrap.php';

if (is_logged_in()) {
    header('Location: cotizador.php');
    exit;
}

$error = '';
$username = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $username = trim($_POST['username'] ?? '');
    $password = $_POST['password'] ?? '';

    if ($username === '' || $password === '') {
        $error = 'Ingresa usuario y contraseña.';
    } elseif (!attempt_login($username, $password, $users)) {
        $error = 'Credenciales inválidas.';
    } else {
        header('Location: cotizador.php');
        exit;
    }
}

render_header('Acceso al cotizador');
?>
<section class="panel">
    <p>Inicia sesión para acceder a la base inicial del cotizador.</p>

    <?php if ($error !== ''): ?>
        <div class="message message-error"><?php echo htmlspecialchars($error, ENT_QUOTES, 'UTF-8'); ?></div>
    <?php endif; ?>

    <form method="post" action="login.php" class="form-grid">
        <div>
            <label for="username">Usuario</label>
            <input id="username" name="username" type="text" value="<?php echo htmlspecialchars($username, ENT_QUOTES, 'UTF-8'); ?>" autocomplete="username" required>
        </div>
        <div>
            <label for="password">Contraseña</label>
            <input id="password" name="password" type="password" autocomplete="current-password" required>
        </div>
        <div class="actions">
            <button type="submit" class="button">Entrar</button>
        </div>
    </form>

    <div class="credentials">
        <strong>Usuarios iniciales:</strong>
        <ul>
            <li>admin</li>
            <li>ventas</li>
        </ul>
        <p class="help-text">Las contraseñas se validan solo del lado servidor y la estructura queda lista para migrar usuarios a base de datos más adelante.</p>
    </div>
</section>
<?php render_footer(); ?>
