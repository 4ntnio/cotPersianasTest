<?php

function round_up_value($value, $precision)
{
    $factor = pow(10, (int) $precision);

    return ceil(((float) $value) * $factor) / $factor;
}

function format_money($value)
{
    return number_format((float) $value, 2);
}

function get_quote_items()
{
    return $_SESSION['quote_items'] ?? [];
}

function add_quote_item(array $item)
{
    if (!isset($_SESSION['quote_items'])) {
        $_SESSION['quote_items'] = [];
    }

    $_SESSION['quote_items'][] = $item;
}

function get_quote_summary(array $items)
{
    $total = 0.00;

    foreach ($items as $item) {
        $total += $item['total_price'];
    }

    return [
        'count' => count($items),
        'total' => round_up_value($total, 2),
    ];
}

function build_quote_from_input(array $input, array $catalog)
{
    $rules = $catalog['measurement_rules'];
    $precision = $rules['precision'];
    $errors = [];

    $type = trim($input['tipo'] ?? '');
    $modelName = trim($input['modelo'] ?? '');
    $operation = trim($input['accionamiento'] ?? '');
    $motorCode = trim($input['motor'] ?? '');
    $controlCode = trim($input['control'] ?? 'none');
    $widthInput = trim($input['ancho'] ?? '');
    $heightInput = trim($input['alto'] ?? '');

    if ($type === '' || !in_array($type, get_blind_types($catalog), true)) {
        $errors[] = 'Selecciona un tipo de persiana válido.';
    }

    $model = $type !== '' ? find_model_for_type($catalog, $type, $modelName) : null;

    if (!$model) {
        $errors[] = 'Selecciona una tela/modelo compatible.';
    }

    if ($operation === '' || !in_array($operation, get_operation_modes($catalog), true)) {
        $errors[] = 'Selecciona un accionamiento válido.';
    }

    if ($widthInput === '' || $heightInput === '') {
        $errors[] = 'Captura ancho y alto.';
    } elseif (!is_numeric($widthInput) || !is_numeric($heightInput)) {
        $errors[] = 'Ancho y alto deben ser valores numéricos.';
    }

    $width = is_numeric($widthInput) ? round_up_value($widthInput, $precision) : null;
    $height = is_numeric($heightInput) ? round_up_value($heightInput, $precision) : null;

    if ($width !== null && $height !== null) {
        if ($width < $rules['min_capture_width'] || $height < $rules['min_capture_height']) {
            $errors[] = 'El ancho y el alto mínimos capturables son 0.20 m.';
        }

        $maxWidth = $rules['max_capture_width'];
        $maxHeight = $rules['max_capture_height'];

        if ($type === 'Pertina') {
            $maxWidth = 5.00;
            $maxHeight = 4.00;
        }

        if ($width > $maxWidth) {
            if ($model && $width > $model['max_width']) {
                $errors[] = 'El ancho insertado supera el ancho de rollo. Analiza alternativas para el cliente.';
            } else {
                $errors[] = 'El ancho excede el máximo permitido para esta fase.';
            }
        }

        if ($height > $maxHeight) {
            $errors[] = 'El alto excede el máximo permitido para esta fase.';
        }

        if ($model && $width > $model['max_width']) {
            $errors[] = 'El ancho insertado supera el ancho de rollo. Analiza alternativas para el cliente.';
        }

        if ($model && $model['max_height'] !== null && $height > $model['max_height']) {
            $errors[] = 'El alto insertado supera el alto máximo permitido para este modelo.';
        }
    }

    $motor = null;
    $control = null;

    if ($operation === 'Motorizado') {
        $motor = find_motor($catalog, $type, $motorCode);

        if (!$motor) {
            $errors[] = 'Selecciona un motor compatible.';
        }

        if ($motor && $width !== null && $height !== null) {
            if ($width < $motor['min_width']) {
                $errors[] = 'La medida capturada no alcanza el mínimo requerido para el motor seleccionado.';
            }

            $profileMatch = false;

            foreach ($motor['profiles'] as $profile) {
                if ($width <= $profile['max_width'] && $height <= $profile['max_height']) {
                    $profileMatch = true;
                    break;
                }
            }

            if (!$profileMatch) {
                $errors[] = 'Las medidas no son compatibles con el motor seleccionado.';
            }
        }

        if ($motor) {
            $control = find_control($catalog, $motor, $controlCode === '' ? 'none' : $controlCode);

            if (!$control) {
                $errors[] = 'Selecciona un control remoto válido.';
            }
        }
    }

    $errors = array_values(array_unique($errors));

    if (!empty($errors) || !$model || $width === null || $height === null) {
        return [
            'errors' => $errors,
            'quote' => null,
        ];
    }

    $billableWidth = max($width, $rules['min_billable_width']);
    $billableHeight = max($height, $rules['min_billable_height']);
    $area = round_up_value($billableWidth * $billableHeight, $precision);
    $basePrice = round_up_value($area * $model['price'], $precision);
    $motorPrice = $motor ? round_up_value($motor['price'], $precision) : 0.00;
    $controlPrice = $control ? round_up_value($control['price'], $precision) : 0.00;
    $totalPrice = round_up_value($basePrice + $motorPrice + $controlPrice, $precision);

    return [
        'errors' => [],
        'quote' => [
            'type' => $type,
            'model' => $model['name'],
            'color' => 'Por definir',
            'operation' => $operation,
            'width' => $width,
            'height' => $height,
            'billable_width' => $billableWidth,
            'billable_height' => $billableHeight,
            'area' => $area,
            'price_per_m2' => round_up_value($model['price'], $precision),
            'base_price' => $basePrice,
            'motor_name' => $motor ? $motor['name'] : 'No aplica',
            'motor_price' => $motorPrice,
            'control_name' => $control ? $control['name'] : 'No aplica',
            'control_price' => $controlPrice,
            'total_price' => $totalPrice,
        ],
    ];
}
