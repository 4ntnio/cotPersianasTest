<?php

return [
    'types' => [
        'Enrollables',
        'Sheer elegance',
        'Panel',
        'Romana',
        'Shangri-la',
        'Pertina',
    ],
    'measurement_rules' => [
        'min_capture_width' => 0.20,
        'min_capture_height' => 0.20,
        'max_capture_width' => 3.50,
        'max_capture_height' => 3.50,
        'min_billable_width' => 1.00,
        'min_billable_height' => 1.00,
        'precision' => 2,
    ],
    'type_notes' => [
        'Shangri-la' => 'No se garantiza el empate de lienzos lado a lado.',
    ],
    'interior_fabrics' => [
        [
            'name' => 'Pirita',
            'max_width' => 2.00,
            'prices' => [
                'Enrollables' => 1056.00,
                'Panel' => 1215.00,
                'Romana' => 2084.00,
            ],
        ],
        [
            'name' => 'Inspiración',
            'max_width' => 2.50,
            'prices' => [
                'Enrollables' => 1351.00,
                'Panel' => 1545.00,
                'Romana' => 2310.00,
            ],
        ],
        [
            'name' => 'Zircón',
            'max_width' => 2.60,
            'prices' => [
                'Enrollables' => 1085.00,
                'Panel' => 1349.00,
                'Romana' => 2051.00,
            ],
        ],
        [
            'name' => 'Onix',
            'max_width' => 3.00,
            'prices' => [
                'Enrollables' => 1450.00,
                'Panel' => 1578.00,
                'Romana' => 2640.00,
            ],
        ],
        [
            'name' => 'Raccon',
            'max_width' => 3.00,
            'prices' => [
                'Enrollables' => 1219.00,
                'Panel' => 1518.00,
                'Romana' => 2163.00,
            ],
        ],
        [
            'name' => 'Screen 2000/4000',
            'max_width' => 3.00,
            'prices' => [
                'Enrollables' => 979.00,
                'Panel' => 1591.00,
                'Romana' => 2526.00,
            ],
        ],
        [
            'name' => 'Hemp',
            'max_width' => 2.80,
            'prices' => [
                'Enrollables' => 1195.00,
                'Panel' => 1560.00,
                'Romana' => 2161.00,
            ],
        ],
        [
            'name' => 'Petunia',
            'max_width' => 3.00,
            'prices' => [
                'Enrollables' => 977.00,
                'Romana' => 2225.00,
            ],
        ],
        [
            'name' => 'Driftwood',
            'max_width' => 3.00,
            'prices' => [
                'Enrollables' => 1824.00,
                'Romana' => 3164.00,
            ],
        ],
        [
            'name' => 'Tamina / Tamina II',
            'max_width' => 2.50,
            'prices' => [
                'Enrollables' => 1175.00,
                'Romana' => 2341.00,
            ],
        ],
        [
            'name' => 'Pandora',
            'max_width' => 3.00,
            'prices' => [
                'Enrollables' => 1587.00,
                'Romana' => 2467.00,
            ],
        ],
        [
            'name' => 'Diamond Blackout',
            'max_width' => 3.00,
            'prices' => [
                'Enrollables' => 1804.00,
                'Romana' => 2889.00,
            ],
        ],
        [
            'name' => 'Gaemi',
            'max_width' => 3.00,
            'prices' => [
                'Enrollables' => 1644.00,
                'Romana' => 2691.00,
            ],
        ],
        [
            'name' => 'Screen black out',
            'max_width' => 2.50,
            'prices' => [
                'Enrollables' => 1285.00,
                'Romana' => 2352.00,
            ],
        ],
        [
            'name' => 'Melania',
            'max_width' => 3.00,
            'prices' => [
                'Enrollables' => 1263.00,
                'Romana' => 2218.00,
            ],
        ],
    ],
    'special_types' => [
        'Sheer elegance' => [
            'models' => [
                ['name' => 'Regular', 'price' => 1303.00, 'max_width' => 2.75],
                ['name' => 'Wood look', 'price' => 1582.00, 'max_width' => 2.75],
                ['name' => 'Lien', 'price' => 1851.00, 'max_width' => 2.75],
                ['name' => 'Hera', 'price' => 2594.00, 'max_width' => 2.78],
                ['name' => 'Mia', 'price' => 2764.00, 'max_width' => 2.78],
                ['name' => 'Maple', 'price' => 2931.00, 'max_width' => 2.78],
                ['name' => 'Antares', 'price' => 3100.00, 'max_width' => 2.75],
                ['name' => 'Mitra', 'price' => 2616.00, 'max_width' => 2.75],
                ['name' => 'Winwood', 'price' => 3133.00, 'max_width' => 2.78],
                ['name' => 'Vella', 'price' => 2192.00, 'max_width' => 2.78],
                ['name' => 'Zinerva', 'price' => 2687.00, 'max_width' => 2.50],
            ],
        ],
        'Shangri-la' => [
            'models' => [
                ['name' => 'Azara', 'price' => 3314.00, 'max_width' => 2.60],
                ['name' => 'Sadira', 'price' => 3773.00, 'max_width' => 2.60],
            ],
        ],
        'Pertina' => [
            'models' => [
                ['name' => 'Pertina Traslúcida', 'price' => 1500.00, 'max_width' => 5.00, 'max_height' => 4.00],
                ['name' => 'Pertina Blackout', 'price' => 1900.00, 'max_width' => 5.00, 'max_height' => 4.00],
            ],
        ],
    ],
    'operation_modes' => [
        'Manual',
        'Motorizado',
    ],
    'motors' => [
        [
            'code' => 'altus-28-15-rts-wf-t38',
            'name' => 'Motor Somfy Altus 28-1.5 RTS WF T38 (Baterías)',
            'price' => 6699.00,
            'min_width' => 0.75,
            'compatible_types' => ['Enrollables', 'Sheer elegance', 'Shangri-la'],
            'remote_system' => 'RTS',
            'profiles' => [
                ['max_width' => 2.00, 'max_height' => 3.00, 'detail' => 'hasta 2.00 x 3.00 con 1 lienzo'],
            ],
        ],
        [
            'code' => 'altus-28-15-rts-wf-t43',
            'name' => 'Motor Somfy Altus 28-1.5 RTS WF T43 (Baterías)',
            'price' => 6699.00,
            'min_width' => 0.75,
            'compatible_types' => ['Enrollables', 'Sheer elegance', 'Shangri-la'],
            'remote_system' => 'RTS',
            'profiles' => [
                ['max_width' => 2.50, 'max_height' => 3.00, 'detail' => 'hasta 2.50 x 3.00 con facia 1 lienzo'],
            ],
        ],
        [
            'code' => 'lsn-406-rts',
            'name' => 'Motor Somfy LSN 406 RTS',
            'price' => 5298.00,
            'min_width' => 0.75,
            'compatible_types' => ['Enrollables', 'Sheer elegance', 'Shangri-la'],
            'remote_system' => 'RTS',
            'profiles' => [
                ['max_width' => 2.80, 'max_height' => 3.00, 'detail' => 'hasta 2.80 x 3.00 con facia 1 lienzo'],
                ['max_width' => 2.80, 'max_height' => 3.50, 'detail' => 'hasta 2.80 x 3.50 sin facia 1 lienzo'],
            ],
        ],
        [
            'code' => 'lsn-406-rts-2l',
            'name' => 'Motor Somfy LSN 406 RTS 2L',
            'price' => 6099.00,
            'min_width' => 0.75,
            'compatible_types' => ['Enrollables', 'Sheer elegance', 'Shangri-la'],
            'remote_system' => 'RTS',
            'profiles' => [
                ['max_width' => 4.50, 'max_height' => 3.00, 'detail' => 'hasta 4.50 x 3.00 con facia 2 lienzos'],
                ['max_width' => 4.50, 'max_height' => 3.50, 'detail' => 'hasta 4.50 x 3.50 sin facia 2 lienzos'],
            ],
        ],
        [
            'code' => 'altus-510-rts',
            'name' => 'Motor Somfy Altus 510 RTS',
            'price' => 7837.00,
            'min_width' => 1.00,
            'compatible_types' => ['Enrollables'],
            'remote_system' => 'RTS',
            'profiles' => [
                ['max_width' => 4.20, 'max_height' => 3.50, 'detail' => 'hasta 4.20 x 3.50 con 1 lienzo'],
            ],
        ],
        [
            'code' => 'altus-510-rts-2l',
            'name' => 'Motor Somfy Altus 510 RTS 2L',
            'price' => 8789.00,
            'min_width' => 1.00,
            'compatible_types' => ['Enrollables'],
            'remote_system' => 'RTS',
            'profiles' => [
                ['max_width' => 8.40, 'max_height' => 3.50, 'detail' => 'hasta 8.40 x 3.50 con 2 lienzos'],
            ],
        ],
        [
            'code' => 'altus-510-rts-3l',
            'name' => 'Motor Somfy Altus 510 RTS 3L',
            'price' => 9845.00,
            'min_width' => 1.00,
            'compatible_types' => ['Enrollables'],
            'remote_system' => 'RTS',
            'profiles' => [
                ['max_width' => 9.00, 'max_height' => 3.00, 'detail' => 'hasta 9.00 x 3.00 con 3 lienzos'],
            ],
        ],
    ],
    'controls' => [
        ['code' => 'none', 'name' => 'Sin control remoto', 'price' => 0.00],
        ['code' => 'situo-1-rts-pure', 'name' => 'Control Somfy Situo 1 RTS Pure', 'price' => 1241.00],
        ['code' => 'situo-2-rts-pure', 'name' => 'Control Somfy Situo 2 RTS Pure', 'price' => 1881.00],
        ['code' => 'situo-5-rts-pure', 'name' => 'Control Somfy Situo 5 RTS Pure', 'price' => 2306.00],
        ['code' => 'situo-5-var-rts-pure', 'name' => 'Control Somfy Situo 5 VAR RTS Pure', 'price' => 4147.00],
        ['code' => 'telis-6-rts-timer-pure', 'name' => 'Control Somfy Telis 6 RTS Timer Pure', 'price' => 8090.00],
        ['code' => 'telis-16-rts-pure', 'name' => 'Control Somfy Telis 16 RTS Pure', 'price' => 7168.00],
        ['code' => 'smoove-a-pared-1-rts', 'name' => 'Control Somfy Smoove a Pared 1 RTS', 'price' => 1798.00],
    ],
];
