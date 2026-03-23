<?php

function get_quote_counter_path()
{
    return __DIR__ . '/../storage/quote-counter.txt';
}

function ensure_quote_counter_exists($path)
{
    if (file_exists($path)) {
        return;
    }

    file_put_contents($path, "0\n", LOCK_EX);
}

function next_quote_folio($path)
{
    ensure_quote_counter_exists($path);

    $handle = fopen($path, 'c+');

    if (!$handle) {
        throw new RuntimeException('No fue posible abrir el archivo del contador de cotizaciones.');
    }

    if (!flock($handle, LOCK_EX)) {
        fclose($handle);
        throw new RuntimeException('No fue posible bloquear el archivo del contador de cotizaciones.');
    }

    $currentValue = trim(stream_get_contents($handle));
    $currentNumber = ctype_digit($currentValue) ? (int) $currentValue : 0;
    $nextNumber = $currentNumber + 1;

    rewind($handle);
    ftruncate($handle, 0);
    fwrite($handle, (string) $nextNumber . PHP_EOL);
    fflush($handle);
    flock($handle, LOCK_UN);
    fclose($handle);

    return sprintf('COT-%04d', $nextNumber);
}
