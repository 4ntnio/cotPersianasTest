<?php

require __DIR__ . '/includes/bootstrap.php';

require_login();

$types = get_blind_types($catalog);
$operationModes = get_operation_modes($catalog);
$items = get_quote_items();
$summary = get_quote_summary($items);

$formData = [
    'tipo' => '',
    'modelo' => '',
    'ancho' => '',
    'alto' => '',
    'accionamiento' => 'Manual',
    'motor' => '',
    'control' => 'none',
];

$messages = [
    'errors' => [],
    'success' => '',
    'info' => '',
];
$currentQuote = null;
$action = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $formData['tipo'] = trim($_POST['tipo'] ?? '');
    $formData['modelo'] = trim($_POST['modelo'] ?? '');
    $formData['ancho'] = trim($_POST['ancho'] ?? '');
    $formData['alto'] = trim($_POST['alto'] ?? '');
    $formData['accionamiento'] = trim($_POST['accionamiento'] ?? 'Manual');
    $formData['motor'] = trim($_POST['motor'] ?? '');
    $formData['control'] = trim($_POST['control'] ?? 'none');
    $action = trim($_POST['form_action'] ?? '');

    if ($action === 'refresh') {
        $messages['info'] = 'Se actualizaron las opciones compatibles según la selección actual.';
    } elseif ($action === 'quote_single' || $action === 'add_piece') {
        $quoteResult = build_quote_from_input($formData, $catalog);
        $messages['errors'] = $quoteResult['errors'];
        $currentQuote = $quoteResult['quote'];

        if ($currentQuote && $action === 'quote_single') {
            $messages['success'] = 'Cotización individual calculada correctamente.';
        }

        if ($currentQuote && $action === 'add_piece') {
            add_quote_item($currentQuote);
            $items = get_quote_items();
            $summary = get_quote_summary($items);
            $messages['success'] = 'La persiana se agregó a la tabla temporal.';
        }
    } elseif ($action === 'quote_all') {
        if ($summary['count'] <= 2) {
            $messages['errors'][] = 'Agrega al menos tres persianas antes de usar "Cotizar PERSIANAS".';
        } else {
            $messages['success'] = 'Cotización acumulada generada con las persianas agregadas.';
        }
    }
}

$models = get_models_for_type($catalog, $formData['tipo']);
if ($formData['modelo'] !== '' && !isset($models[$formData['modelo']])) {
    $formData['modelo'] = '';
}

$motors = get_motors_for_type($catalog, $formData['tipo']);
if ($formData['motor'] !== '' && !isset($motors[$formData['motor']])) {
    $formData['motor'] = '';
}

$selectedMotor = $formData['motor'] !== '' ? ($motors[$formData['motor']] ?? null) : null;
$controls = $formData['accionamiento'] === 'Motorizado' && !empty($motors) ? get_controls_for_motor($catalog, $selectedMotor) : [];
if ($formData['control'] !== '' && $formData['control'] !== 'none' && !isset($controls[$formData['control']])) {
    $formData['control'] = 'none';
}

$typeNote = get_type_note($catalog, $formData['tipo']);
$disableQuoteAll = $summary['count'] <= 2;

render_header('Cotizador de persianas - Fase 2');
?>
<section class="panel panel-spacing">
    <div class="section-heading">
        <div>
            <p class="eyebrow">Fase 2</p>
            <h2>Cotización de persianas</h2>
        </div>
        <div class="summary-chip-group">
            <span class="summary-chip">Piezas agregadas: <?php echo htmlspecialchars((string) $summary['count'], ENT_QUOTES, 'UTF-8'); ?></span>
            <span class="summary-chip">Total acumulado: $<?php echo htmlspecialchars(format_money($summary['total']), ENT_QUOTES, 'UTF-8'); ?></span>
        </div>
    </div>

    <p>Selecciona el tipo de persiana, la tela/modelo compatible y el accionamiento. El cálculo final se resuelve siempre en backend.</p>

    <?php foreach ($messages['errors'] as $error): ?>
        <div class="message message-error"><?php echo htmlspecialchars($error, ENT_QUOTES, 'UTF-8'); ?></div>
    <?php endforeach; ?>

    <?php if ($messages['success'] !== ''): ?>
        <div class="message message-success"><?php echo htmlspecialchars($messages['success'], ENT_QUOTES, 'UTF-8'); ?></div>
    <?php endif; ?>

    <?php if ($messages['info'] !== ''): ?>
        <div class="message message-info"><?php echo htmlspecialchars($messages['info'], ENT_QUOTES, 'UTF-8'); ?></div>
    <?php endif; ?>

    <?php if ($typeNote !== ''): ?>
        <div class="message message-warning"><?php echo htmlspecialchars($typeNote, ENT_QUOTES, 'UTF-8'); ?></div>
    <?php endif; ?>

    <form method="post" action="cotizador.php" class="form-grid">
        <input type="hidden" name="form_action" value="quote_single" id="form_action">

        <div class="form-grid columns-2 columns-3">
            <div>
                <label for="tipo">Tipo de persiana</label>
                <select id="tipo" name="tipo" onchange="document.getElementById('form_action').value='refresh'; this.form.submit();">
                    <option value="">Selecciona una opción</option>
                    <?php foreach ($types as $type): ?>
                        <option value="<?php echo htmlspecialchars($type, ENT_QUOTES, 'UTF-8'); ?>" <?php echo $formData['tipo'] === $type ? 'selected' : ''; ?>><?php echo htmlspecialchars($type, ENT_QUOTES, 'UTF-8'); ?></option>
                    <?php endforeach; ?>
                </select>
                <noscript><button type="submit" class="button button-secondary inline-button" onclick="document.getElementById('form_action').value='refresh';">Actualizar modelos</button></noscript>
            </div>
            <div>
                <label for="modelo">Tela / modelo</label>
                <select id="modelo" name="modelo">
                    <option value="">Selecciona una opción</option>
                    <?php foreach ($models as $model): ?>
                        <option value="<?php echo htmlspecialchars($model['name'], ENT_QUOTES, 'UTF-8'); ?>" <?php echo $formData['modelo'] === $model['name'] ? 'selected' : ''; ?>><?php echo htmlspecialchars($model['name'], ENT_QUOTES, 'UTF-8'); ?></option>
                    <?php endforeach; ?>
                </select>
            </div>
            <div>
                <label for="accionamiento">Accionamiento</label>
                <select id="accionamiento" name="accionamiento" onchange="document.getElementById('form_action').value='refresh'; this.form.submit();">
                    <?php foreach ($operationModes as $operationMode): ?>
                        <option value="<?php echo htmlspecialchars($operationMode, ENT_QUOTES, 'UTF-8'); ?>" <?php echo $formData['accionamiento'] === $operationMode ? 'selected' : ''; ?>><?php echo htmlspecialchars($operationMode, ENT_QUOTES, 'UTF-8'); ?></option>
                    <?php endforeach; ?>
                </select>
                <noscript><button type="submit" class="button button-secondary inline-button" onclick="document.getElementById('form_action').value='refresh';">Actualizar accionamiento</button></noscript>
            </div>
        </div>

        <div class="form-grid columns-2 columns-3">
            <div>
                <label for="ancho">Ancho (m)</label>
                <input id="ancho" name="ancho" type="number" min="0.20" step="0.01" value="<?php echo htmlspecialchars($formData['ancho'], ENT_QUOTES, 'UTF-8'); ?>" required>
                <p class="help-text">Mínimo capturable: 0.20 m. Mínimo cobrable: 1.00 m.</p>
            </div>
            <div>
                <label for="alto">Alto (m)</label>
                <input id="alto" name="alto" type="number" min="0.20" step="0.01" value="<?php echo htmlspecialchars($formData['alto'], ENT_QUOTES, 'UTF-8'); ?>" required>
                <p class="help-text">Máximo general: 3.50 m. En Pertina el alto máximo es 4.00 m.</p>
            </div>
            <div>
                <label>Color</label>
                <div class="static-field">Por definir</div>
                <p class="help-text">No afecta el precio en esta fase.</p>
            </div>
        </div>

        <?php if ($formData['accionamiento'] === 'Motorizado'): ?>
            <?php if (empty($motors)): ?>
                <div class="message message-warning">No hay motores compatibles para el tipo de persiana seleccionado en esta fase.</div>
            <?php else: ?>
                <div class="form-grid columns-2">
                    <div>
                        <label for="motor">Motor compatible</label>
                        <select id="motor" name="motor">
                            <option value="">Selecciona una opción</option>
                            <?php foreach ($motors as $motor): ?>
                                <option value="<?php echo htmlspecialchars($motor['code'], ENT_QUOTES, 'UTF-8'); ?>" <?php echo $formData['motor'] === $motor['code'] ? 'selected' : ''; ?>><?php echo htmlspecialchars($motor['name'], ENT_QUOTES, 'UTF-8'); ?></option>
                            <?php endforeach; ?>
                        </select>
                        <p class="help-text">Solo se muestran motores compatibles con el tipo seleccionado.</p>
                    </div>
                    <div>
                        <label for="control">Control remoto</label>
                        <select id="control" name="control">
                            <?php foreach ($controls as $control): ?>
                                <option value="<?php echo htmlspecialchars($control['code'], ENT_QUOTES, 'UTF-8'); ?>" <?php echo $formData['control'] === $control['code'] ? 'selected' : ''; ?>><?php echo htmlspecialchars($control['name'], ENT_QUOTES, 'UTF-8'); ?></option>
                            <?php endforeach; ?>
                        </select>
                        <p class="help-text">Disponible solo para motores RTS. Puedes elegir "Sin control remoto".</p>
                    </div>
                </div>
            <?php endif; ?>
        <?php else: ?>
            <div class="message message-info">Accionamiento manual seleccionado. No se agrega costo extra en esta fase.</div>
        <?php endif; ?>

        <div class="actions actions-wrap">
            <button type="submit" class="button" onclick="document.getElementById('form_action').value='quote_single';">Cotizar persiana</button>
            <button type="submit" class="button button-secondary" onclick="document.getElementById('form_action').value='add_piece';">Agregar persiana</button>
            <button type="submit" class="button button-dark" onclick="document.getElementById('form_action').value='quote_all';" <?php echo $disableQuoteAll ? 'disabled' : ''; ?>>Cotizar PERSIANAS</button>
        </div>
    </form>
</section>

<?php if ($currentQuote): ?>
    <section class="panel panel-spacing">
        <div class="section-heading">
            <div>
                <p class="eyebrow">Resultado</p>
                <h2>Cotización individual</h2>
            </div>
            <span class="summary-chip">Precio pieza: $<?php echo htmlspecialchars(format_money($currentQuote['total_price']), ENT_QUOTES, 'UTF-8'); ?></span>
        </div>
        <div class="quote-breakdown">
            <div class="breakdown-item"><strong>Tipo:</strong> <?php echo htmlspecialchars($currentQuote['type'], ENT_QUOTES, 'UTF-8'); ?></div>
            <div class="breakdown-item"><strong>Tela / modelo:</strong> <?php echo htmlspecialchars($currentQuote['model'], ENT_QUOTES, 'UTF-8'); ?></div>
            <div class="breakdown-item"><strong>Color:</strong> <?php echo htmlspecialchars($currentQuote['color'], ENT_QUOTES, 'UTF-8'); ?></div>
            <div class="breakdown-item"><strong>Accionamiento:</strong> <?php echo htmlspecialchars($currentQuote['operation'], ENT_QUOTES, 'UTF-8'); ?></div>
            <div class="breakdown-item"><strong>Medidas capturadas:</strong> <?php echo htmlspecialchars(format_money($currentQuote['width']), ENT_QUOTES, 'UTF-8'); ?> × <?php echo htmlspecialchars(format_money($currentQuote['height']), ENT_QUOTES, 'UTF-8'); ?> m</div>
            <div class="breakdown-item"><strong>Medidas cobrables:</strong> <?php echo htmlspecialchars(format_money($currentQuote['billable_width']), ENT_QUOTES, 'UTF-8'); ?> × <?php echo htmlspecialchars(format_money($currentQuote['billable_height']), ENT_QUOTES, 'UTF-8'); ?> m</div>
            <div class="breakdown-item"><strong>Área cobrable:</strong> <?php echo htmlspecialchars(format_money($currentQuote['area']), ENT_QUOTES, 'UTF-8'); ?> m²</div>
            <div class="breakdown-item"><strong>Precio por m²:</strong> $<?php echo htmlspecialchars(format_money($currentQuote['price_per_m2']), ENT_QUOTES, 'UTF-8'); ?></div>
            <div class="breakdown-item"><strong>Subtotal persiana:</strong> $<?php echo htmlspecialchars(format_money($currentQuote['base_price']), ENT_QUOTES, 'UTF-8'); ?></div>
            <div class="breakdown-item"><strong>Motor:</strong> <?php echo htmlspecialchars($currentQuote['motor_name'], ENT_QUOTES, 'UTF-8'); ?></div>
            <div class="breakdown-item"><strong>Control remoto:</strong> <?php echo htmlspecialchars($currentQuote['control_name'], ENT_QUOTES, 'UTF-8'); ?></div>
            <div class="breakdown-item"><strong>Total:</strong> $<?php echo htmlspecialchars(format_money($currentQuote['total_price']), ENT_QUOTES, 'UTF-8'); ?></div>
        </div>
    </section>
<?php endif; ?>

<section class="panel panel-spacing">
    <div class="section-heading">
        <div>
            <p class="eyebrow">Tabla temporal</p>
            <h2>Piezas acumuladas</h2>
        </div>
        <div class="summary-chip-group">
            <span class="summary-chip">Cantidad: <?php echo htmlspecialchars((string) $summary['count'], ENT_QUOTES, 'UTF-8'); ?></span>
            <span class="summary-chip">Total acumulado: $<?php echo htmlspecialchars(format_money($summary['total']), ENT_QUOTES, 'UTF-8'); ?></span>
        </div>
    </div>

    <?php if (empty($items)): ?>
        <p>Aún no hay persianas agregadas a la tabla temporal.</p>
    <?php else: ?>
        <div class="table-wrapper">
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Tipo</th>
                        <th>Tela / modelo</th>
                        <th>Color</th>
                        <th>Accionamiento</th>
                        <th>Precio pieza</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($items as $index => $item): ?>
                        <tr>
                            <td><?php echo htmlspecialchars((string) ($index + 1), ENT_QUOTES, 'UTF-8'); ?></td>
                            <td><?php echo htmlspecialchars($item['type'], ENT_QUOTES, 'UTF-8'); ?></td>
                            <td><?php echo htmlspecialchars($item['model'], ENT_QUOTES, 'UTF-8'); ?></td>
                            <td><?php echo htmlspecialchars($item['color'], ENT_QUOTES, 'UTF-8'); ?></td>
                            <td><?php echo htmlspecialchars($item['operation'], ENT_QUOTES, 'UTF-8'); ?></td>
                            <td>$<?php echo htmlspecialchars(format_money($item['total_price']), ENT_QUOTES, 'UTF-8'); ?></td>
                        </tr>
                    <?php endforeach; ?>
                </tbody>
                <tfoot>
                    <tr>
                        <td colspan="5">Total acumulado</td>
                        <td>$<?php echo htmlspecialchars(format_money($summary['total']), ENT_QUOTES, 'UTF-8'); ?></td>
                    </tr>
                </tfoot>
            </table>
        </div>
    <?php endif; ?>
</section>
<?php render_footer(); ?>
