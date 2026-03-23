<?php

function is_logged_in()
{
    return isset($_SESSION['user']);
}

function current_user()
{
    return $_SESSION['user'] ?? null;
}

function attempt_login($username, $password, array $users)
{
    if (!isset($users[$username])) {
        return false;
    }

    $user = $users[$username];

    if (!password_verify($password, $user['password_hash'])) {
        return false;
    }

    $_SESSION['user'] = [
        'username' => $user['username'],
        'display_name' => $user['display_name'],
    ];

    session_regenerate_id(true);

    return true;
}

function require_login()
{
    if (is_logged_in()) {
        return;
    }

    header('Location: login.php');
    exit;
}

function logout_user()
{
    $_SESSION = [];

    if (ini_get('session.use_cookies')) {
        $params = session_get_cookie_params();
        setcookie(
            session_name(),
            '',
            time() - 42000,
            $params['path'],
            $params['domain'],
            $params['secure'],
            $params['httponly']
        );
    }

    session_destroy();
}
