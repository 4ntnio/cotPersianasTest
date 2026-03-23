<?php

function render_header($title)
{
    $pageTitle = $title;
    require __DIR__ . '/../templates/header.php';
}

function render_footer()
{
    require __DIR__ . '/../templates/footer.php';
}
