<?php

function get_blind_types(array $catalog)
{
    return $catalog['types'];
}

function get_operation_modes(array $catalog)
{
    return $catalog['operation_modes'];
}

function get_type_note(array $catalog, $type)
{
    return $catalog['type_notes'][$type] ?? '';
}

function get_models_for_type(array $catalog, $type)
{
    if ($type === '') {
        return [];
    }

    $models = [];

    foreach ($catalog['interior_fabrics'] as $fabric) {
        if (!isset($fabric['prices'][$type])) {
            continue;
        }

        $models[$fabric['name']] = [
            'name' => $fabric['name'],
            'price' => $fabric['prices'][$type],
            'max_width' => $fabric['max_width'],
            'max_height' => null,
            'type' => $type,
        ];
    }

    if (isset($catalog['special_types'][$type])) {
        foreach ($catalog['special_types'][$type]['models'] as $model) {
            $models[$model['name']] = [
                'name' => $model['name'],
                'price' => $model['price'],
                'max_width' => $model['max_width'],
                'max_height' => $model['max_height'] ?? null,
                'type' => $type,
            ];
        }
    }

    ksort($models);

    return $models;
}

function find_model_for_type(array $catalog, $type, $modelName)
{
    $models = get_models_for_type($catalog, $type);

    return $models[$modelName] ?? null;
}

function get_motors_for_type(array $catalog, $type)
{
    $motors = [];

    foreach ($catalog['motors'] as $motor) {
        if (!in_array($type, $motor['compatible_types'], true)) {
            continue;
        }

        $motors[$motor['code']] = $motor;
    }

    return $motors;
}

function find_motor(array $catalog, $type, $code)
{
    $motors = get_motors_for_type($catalog, $type);

    return $motors[$code] ?? null;
}

function get_controls_for_motor(array $catalog, $motor)
{
    if ($motor && ($motor['remote_system'] ?? '') !== 'RTS') {
        return [];
    }

    $controls = [];

    foreach ($catalog['controls'] as $control) {
        $controls[$control['code']] = $control;
    }

    return $controls;
}

function find_control(array $catalog, $motor, $code)
{
    $controls = get_controls_for_motor($catalog, $motor);

    return $controls[$code] ?? null;
}
