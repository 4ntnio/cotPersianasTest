<?php

function get_quote_storage_directory()
{
    return __DIR__ . '/../storage';
}

function build_document_meta(array $input)
{
    return [
        'client_name' => sanitize_document_text($input['cliente'] ?? ''),
        'client_phone' => sanitize_document_text($input['telefono'] ?? ''),
        'client_address' => sanitize_document_text($input['direccion'] ?? ''),
        'validity' => sanitize_document_text($input['vigencia'] ?? '7 días naturales'),
        'observations' => sanitize_document_text($input['observaciones'] ?? ''),
    ];
}

function sanitize_document_text($value)
{
    $value = trim((string) $value);
    $value = preg_replace('/\s+/u', ' ', $value);

    return $value;
}

function validate_document_meta(array $meta)
{
    $errors = [];

    if ($meta['client_name'] === '') {
        $errors[] = 'Ingresa el nombre del cliente para generar la cotización.';
    }

    if ($meta['client_phone'] === '') {
        $errors[] = 'Ingresa el teléfono o WhatsApp del cliente para generar la cotización.';
    }

    if ($meta['validity'] === '') {
        $errors[] = 'Ingresa la vigencia de la cotización.';
    }

    return $errors;
}

function escape_docx($value)
{
    return htmlspecialchars((string) $value, ENT_XML1 | ENT_COMPAT, 'UTF-8');
}

function make_docx_paragraph($text, $options = [])
{
    $escapedText = escape_docx($text);
    $paragraphStyle = '';

    if (!empty($options['spacing_before']) || !empty($options['spacing_after'])) {
        $paragraphStyle .= '<w:spacing';
        if (!empty($options['spacing_before'])) {
            $paragraphStyle .= ' w:before="' . (int) $options['spacing_before'] . '"';
        }
        if (!empty($options['spacing_after'])) {
            $paragraphStyle .= ' w:after="' . (int) $options['spacing_after'] . '"';
        }
        $paragraphStyle .= '/>';
    }

    if (!empty($options['align'])) {
        $paragraphStyle .= '<w:jc w:val="' . escape_docx($options['align']) . '"/>';
    }

    if (!empty($options['keep_next'])) {
        $paragraphStyle .= '<w:keepNext/>';
    }

    $runStyle = '';
    if (!empty($options['bold'])) {
        $runStyle .= '<w:b/>';
    }
    if (!empty($options['color'])) {
        $runStyle .= '<w:color w:val="' . escape_docx($options['color']) . '"/>';
    }
    if (!empty($options['size'])) {
        $runStyle .= '<w:sz w:val="' . (int) $options['size'] . '"/><w:szCs w:val="' . (int) $options['size'] . '"/>';
    }

    return '<w:p><w:pPr>' . $paragraphStyle . '</w:pPr><w:r><w:rPr>' . $runStyle . '</w:rPr><w:t xml:space="preserve">' . $escapedText . '</w:t></w:r></w:p>';
}

function make_docx_multiline_paragraph(array $lines, $options = [])
{
    $paragraphStyle = '';

    if (!empty($options['spacing_before']) || !empty($options['spacing_after'])) {
        $paragraphStyle .= '<w:spacing';
        if (!empty($options['spacing_before'])) {
            $paragraphStyle .= ' w:before="' . (int) $options['spacing_before'] . '"';
        }
        if (!empty($options['spacing_after'])) {
            $paragraphStyle .= ' w:after="' . (int) $options['spacing_after'] . '"';
        }
        $paragraphStyle .= '/>';
    }

    if (!empty($options['align'])) {
        $paragraphStyle .= '<w:jc w:val="' . escape_docx($options['align']) . '"/>';
    }

    $runStyle = '';
    if (!empty($options['bold'])) {
        $runStyle .= '<w:b/>';
    }
    if (!empty($options['color'])) {
        $runStyle .= '<w:color w:val="' . escape_docx($options['color']) . '"/>';
    }
    if (!empty($options['size'])) {
        $runStyle .= '<w:sz w:val="' . (int) $options['size'] . '"/><w:szCs w:val="' . (int) $options['size'] . '"/>';
    }

    $chunks = [];
    foreach (array_values($lines) as $index => $line) {
        $chunks[] = '<w:r><w:rPr>' . $runStyle . '</w:rPr><w:t xml:space="preserve">' . escape_docx($line) . '</w:t></w:r>';
        if ($index < count($lines) - 1) {
            $chunks[] = '<w:r><w:br/></w:r>';
        }
    }

    return '<w:p><w:pPr>' . $paragraphStyle . '</w:pPr>' . implode('', $chunks) . '</w:p>';
}

function make_docx_table($rows, $options = [])
{
    $width = isset($options['width']) ? (int) $options['width'] : 9000;
    $borders = $options['borders'] ?? true;
    $cellMargin = $options['cell_margin'] ?? 90;
    $tableStyle = '<w:tblW w:w="' . $width . '" w:type="dxa"/>';
    $tableStyle .= '<w:tblCellMar>'
        . '<w:top w:w="' . (int) $cellMargin . '" w:type="dxa"/>'
        . '<w:left w:w="' . (int) $cellMargin . '" w:type="dxa"/>'
        . '<w:bottom w:w="' . (int) $cellMargin . '" w:type="dxa"/>'
        . '<w:right w:w="' . (int) $cellMargin . '" w:type="dxa"/>'
        . '</w:tblCellMar>';

    if ($borders) {
        $tableStyle .= '<w:tblBorders>'
            . '<w:top w:val="single" w:sz="6" w:space="0" w:color="D9E2EC"/>'
            . '<w:left w:val="single" w:sz="6" w:space="0" w:color="D9E2EC"/>'
            . '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="D9E2EC"/>'
            . '<w:right w:val="single" w:sz="6" w:space="0" w:color="D9E2EC"/>'
            . '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E5E7EB"/>'
            . '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="E5E7EB"/>'
            . '</w:tblBorders>';
    }

    return '<w:tbl><w:tblPr>' . $tableStyle . '</w:tblPr>' . implode('', $rows) . '</w:tbl>';
}

function make_docx_table_cell($contentXml, $options = [])
{
    $width = isset($options['width']) ? (int) $options['width'] : 1800;
    $background = !empty($options['background']) ? '<w:shd w:val="clear" w:color="auto" w:fill="' . escape_docx($options['background']) . '"/>' : '';
    $vAlign = !empty($options['valign']) ? '<w:vAlign w:val="' . escape_docx($options['valign']) . '"/>' : '<w:vAlign w:val="top"/>';
    $gridSpan = !empty($options['grid_span']) ? '<w:gridSpan w:val="' . (int) $options['grid_span'] . '"/>' : '';

    return '<w:tc><w:tcPr><w:tcW w:w="' . $width . '" w:type="dxa"/>' . $gridSpan . $background . $vAlign . '</w:tcPr>' . $contentXml . '</w:tc>';
}

function make_docx_table_row($cells, $options = [])
{
    $isHeader = !empty($options['header']);
    $rowXml = '<w:tr>';

    foreach ($cells as $cell) {
        $cellOptions = $cell['options'] ?? [];
        if ($isHeader && empty($cellOptions['background'])) {
            $cellOptions['background'] = '1F2937';
        }
        $rowXml .= make_docx_table_cell($cell['content'], $cellOptions);
    }

    $rowXml .= '</w:tr>';

    return $rowXml;
}

function make_docx_inline_image($relationshipId)
{
    return '<w:p>'
        . '<w:r><w:drawing>'
        . '<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        . '<wp:extent cx="2500000" cy="700000"/>'
        . '<wp:docPr id="1" name="ZAME Logo"/>'
        . '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        . '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        . '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        . '<pic:nvPicPr><pic:cNvPr id="0" name="Logo_zame_2.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        . '<pic:blipFill><a:blip r:embed="' . escape_docx($relationshipId) . '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        . '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2500000" cy="700000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        . '</pic:pic></a:graphicData></a:graphic>'
        . '</wp:inline></w:drawing></w:r></w:p>';
}

function build_quote_item_description(array $item)
{
    $lines = [
        'Suministro y colocación de ' . $item['type'] . '.',
        'Modelo / Tela: ' . $item['model'] . '.',
        'Color: ' . $item['color'] . '.',
        'Accionamiento: ' . $item['operation'] . '.',
    ];

    if ($item['operation'] === 'Motorizado' && $item['motor_name'] !== 'No aplica') {
        $lines[] = 'Incluye motor: ' . $item['motor_name'] . '.';
    }

    if ($item['operation'] === 'Motorizado' && $item['control_name'] !== 'No aplica') {
        $lines[] = 'Incluye control: ' . $item['control_name'] . '.';
    }

    return $lines;
}

function build_quote_docx_xml(array $document)
{
    $meta = $document['meta'];
    $items = $document['items'];
    $company = $document['company'];
    $headerLeft = (!empty($document['logo_relationship_id'])
        ? make_docx_inline_image($document['logo_relationship_id'])
        : make_docx_paragraph($company['name'], ['bold' => true, 'size' => 28, 'spacing_after' => 80]))
        . make_docx_paragraph('Datos del asesor', ['bold' => true, 'color' => '486581', 'spacing_after' => 40])
        . make_docx_paragraph($company['name'], ['bold' => true, 'size' => 22, 'spacing_after' => 10])
        . make_docx_paragraph($company['tagline'], ['color' => '486581', 'spacing_after' => 40])
        . make_docx_paragraph('Correo: zame.motorizacion@gmail.com / ' . $company['email'], ['spacing_after' => 20])
        . make_docx_paragraph('Sitio: ' . $company['website'], ['spacing_after' => 20])
        . make_docx_paragraph('Asesor: Elizabeth Orozco', ['spacing_after' => 20])
        . make_docx_paragraph('WhatsApp Oficina: 477-173-8901');

    $headerRight = implode('', [
        make_docx_paragraph('Datos de la cotización', ['bold' => true, 'color' => '486581', 'align' => 'right', 'spacing_after' => 40]),
        make_docx_paragraph('Folio: ' . $document['folio'], ['bold' => true, 'size' => 24, 'align' => 'right', 'spacing_after' => 30]),
        make_docx_paragraph('Fecha: ' . $document['date'], ['align' => 'right', 'spacing_after' => 20]),
        make_docx_paragraph('Vigencia: ' . $meta['validity'], ['align' => 'right']),
    ]);

    $headerTable = make_docx_table([
        make_docx_table_row([
            ['content' => $headerLeft, 'options' => ['width' => 5800, 'valign' => 'top']],
            ['content' => $headerRight, 'options' => ['width' => 3200, 'valign' => 'top']],
        ]),
    ], ['width' => 9000, 'borders' => false, 'cell_margin' => 40]);

    $clientAddress = $meta['client_address'] !== '' ? $meta['client_address'] : 'Pendiente por compartir.';
    $clientTable = make_docx_table([
        make_docx_table_row([
            ['content' => make_docx_paragraph('Datos del cliente', ['bold' => true, 'color' => 'FFFFFF', 'spacing_after' => 0]), 'options' => ['width' => 9000, 'background' => '334E68', 'grid_span' => 2]],
        ]),
        make_docx_table_row([
            ['content' => make_docx_paragraph('Atención / Cliente: ' . $meta['client_name'], ['spacing_after' => 0]), 'options' => ['width' => 4500]],
            ['content' => make_docx_paragraph('Teléfono / WhatsApp: ' . $meta['client_phone'], ['spacing_after' => 0]), 'options' => ['width' => 4500]],
        ]),
        make_docx_table_row([
            ['content' => make_docx_paragraph('Dirección: ' . $clientAddress, ['spacing_after' => 0]), 'options' => ['width' => 9000, 'grid_span' => 2]],
        ]),
    ], ['width' => 9000]);

    $quoteRows = [];
    $quoteRows[] = make_docx_table_row([
        ['content' => make_docx_paragraph('Cantidad', ['bold' => true, 'color' => 'FFFFFF', 'align' => 'center']), 'options' => ['width' => 1200]],
        ['content' => make_docx_paragraph('Descripción', ['bold' => true, 'color' => 'FFFFFF']), 'options' => ['width' => 5900]],
        ['content' => make_docx_paragraph('Precio', ['bold' => true, 'color' => 'FFFFFF', 'align' => 'center']), 'options' => ['width' => 2200]],
    ], ['header' => true]);

    foreach ($items as $item) {
        $quoteRows[] = make_docx_table_row([
            ['content' => make_docx_paragraph('1', ['align' => 'center', 'spacing_after' => 0]), 'options' => ['width' => 1200, 'valign' => 'center']],
            ['content' => make_docx_multiline_paragraph(build_quote_item_description($item), ['spacing_after' => 0]), 'options' => ['width' => 5600]],
            ['content' => make_docx_paragraph('$' . format_money($item['total_price']), ['align' => 'center', 'bold' => true]), 'options' => ['width' => 2200, 'valign' => 'center']],
        ]);
    }

    $quoteRows[] = make_docx_table_row([
        ['content' => make_docx_paragraph('', ['spacing_after' => 0]), 'options' => ['width' => 1200, 'background' => 'F9FAFB']],
        ['content' => make_docx_paragraph('TOTAL', ['bold' => true, 'align' => 'right']), 'options' => ['width' => 5600, 'background' => 'F9FAFB']],
        ['content' => make_docx_paragraph('$' . format_money($document['total']), ['bold' => true, 'align' => 'center']), 'options' => ['width' => 2200, 'background' => 'F9FAFB']],
    ]);

    $quoteTable = make_docx_table($quoteRows, ['width' => 9000]);

    $paymentLines = [];
    foreach ($company['payment_terms'] as $term) {
        $paymentLines[] = '• ' . $term;
    }

    $importantLines = [];
    foreach ($company['important_notes'] as $note) {
        $importantLines[] = '• ' . $note;
    }

    $conditionsTable = make_docx_table([
        make_docx_table_row([
            ['content' => make_docx_paragraph('Condiciones comerciales', ['bold' => true, 'color' => 'FFFFFF']), 'options' => ['width' => 9000, 'background' => '1F2937']],
        ]),
        make_docx_table_row([
            ['content' => make_docx_paragraph('Formas de pago', ['bold' => true, 'color' => 'C6A15B', 'spacing_after' => 40]) . make_docx_multiline_paragraph($paymentLines), 'options' => ['width' => 9000]],
        ]),
        make_docx_table_row([
            ['content' => make_docx_paragraph('Importante', ['bold' => true, 'color' => 'C6A15B', 'spacing_after' => 40]) . make_docx_multiline_paragraph($importantLines) . make_docx_paragraph('Vigencia comercial: ' . $meta['validity'], ['spacing_before' => 60]) . make_docx_paragraph('Observaciones: ' . ($meta['observations'] !== '' ? $meta['observations'] : 'Sin observaciones adicionales.')), 'options' => ['width' => 9000]],
        ]),
    ], ['width' => 9000]);

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        . 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        . 'xmlns:o="urn:schemas-microsoft-com:office:office" '
        . 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        . 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        . 'xmlns:v="urn:schemas-microsoft-com:vml" '
        . 'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        . 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        . 'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        . 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        . 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        . 'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        . 'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
        . 'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
        . 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        . 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        . 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
        . 'mc:Ignorable="w14 wp14">'
        . '<w:body>'
        . $headerTable
        . make_docx_paragraph('', ['spacing_after' => 80])
        . $clientTable
        . make_docx_paragraph('', ['spacing_after' => 100])
        . make_docx_paragraph($company['intro_text'][0] ?? 'Encontrará a continuación el presupuesto solicitado. Quedamos a sus órdenes para cualquier aclaración.', ['spacing_after' => 200])
        . $quoteTable
        . make_docx_paragraph('', ['spacing_after' => 140])
        . $conditionsTable
        . make_docx_paragraph('', ['spacing_after' => 140])
        . make_docx_paragraph($company['name'] . ' · ' . $company['tagline'], ['bold' => true, 'align' => 'center', 'color' => '1F2937', 'spacing_after' => 40])
        . make_docx_paragraph($company['email'] . ' | ' . $company['website'], ['align' => 'center', 'color' => '486581'])
        . '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1000" w:right="1100" w:bottom="1100" w:left="1100" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
        . '</w:body></w:document>';
}

function build_quote_docx_file(array $document, $outputFile)
{
    if (!class_exists('ZipArchive')) {
        throw new RuntimeException('ZipArchive no está disponible en esta instalación de PHP.');
    }

    $zip = new ZipArchive();

    if ($zip->open($outputFile, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
        throw new RuntimeException('No fue posible crear el archivo DOCX.');
    }

    $createdAt = gmdate('Y-m-d\TH:i:s\Z');
    $coreTitle = escape_docx('Cotización ' . $document['folio']);
    $coreSubject = escape_docx('Cotización comercial ZAME Blinds');
    $coreAuthor = escape_docx($document['company']['name']);
    $logoRelationshipId = null;

    if (!empty($document['company']['logo_path']) && file_exists($document['company']['logo_path'])) {
        $logoRelationshipId = 'rIdLogo';
    }

    $document['logo_relationship_id'] = $logoRelationshipId;
    $documentXml = build_quote_docx_xml($document);

    $zip->addFromString('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        . '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        . '<Default Extension="xml" ContentType="application/xml"/>'
        . '<Default Extension="png" ContentType="image/png"/>'
        . '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        . '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        . '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        . '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        . '</Types>');

    $zip->addFromString('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        . '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        . '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        . '</Relationships>');

    $zip->addFromString('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        . 'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        . '<Application>ZAME Blinds Cotizador</Application>'
        . '</Properties>');

    $zip->addFromString('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        . 'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        . 'xmlns:dcterms="http://purl.org/dc/terms/" '
        . 'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        . 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        . '<dc:title>' . $coreTitle . '</dc:title>'
        . '<dc:subject>' . $coreSubject . '</dc:subject>'
        . '<dc:creator>' . $coreAuthor . '</dc:creator>'
        . '<cp:lastModifiedBy>' . $coreAuthor . '</cp:lastModifiedBy>'
        . '<dcterms:created xsi:type="dcterms:W3CDTF">' . $createdAt . '</dcterms:created>'
        . '<dcterms:modified xsi:type="dcterms:W3CDTF">' . $createdAt . '</dcterms:modified>'
        . '</cp:coreProperties>');

    $zip->addFromString('word/document.xml', $documentXml);
    $zip->addFromString('word/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        . '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/></w:rPr></w:rPrDefault></w:docDefaults>'
        . '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        . '</w:styles>');

    $documentRelationships = [];
    if ($logoRelationshipId !== null) {
        $documentRelationships[] = '<Relationship Id="' . $logoRelationshipId . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/Logo_zame_2.png"/>';
        $zip->addFile($document['company']['logo_path'], 'word/media/Logo_zame_2.png');
    }

    $zip->addFromString('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        . implode('', $documentRelationships)
        . '</Relationships>');

    $zip->close();
}

function output_docx_download(array $items, array $meta, array $company)
{
    $counterPath = get_quote_counter_path();
    $storageDirectory = get_quote_storage_directory();
    $folio = next_quote_folio($counterPath);
    $date = date('d/m/Y');
    $summary = get_quote_summary($items);

    if (!is_dir($storageDirectory)) {
        mkdir($storageDirectory, 0775, true);
    }

    $tempFile = tempnam($storageDirectory, 'docx_');

    if ($tempFile === false) {
        throw new RuntimeException('No fue posible preparar el archivo temporal para la cotización.');
    }

    $docxFile = $tempFile . '.docx';
    @unlink($tempFile);

    build_quote_docx_file([
        'folio' => $folio,
        'date' => $date,
        'meta' => $meta,
        'items' => $items,
        'total' => $summary['total'],
        'company' => $company,
    ], $docxFile);

    header('Content-Description: File Transfer');
    header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    header('Content-Disposition: attachment; filename="' . $folio . '.docx"');
    header('Content-Length: ' . filesize($docxFile));
    header('Cache-Control: no-store, no-cache, must-revalidate');
    header('Pragma: public');

    readfile($docxFile);
    @unlink($docxFile);
    exit;
}
