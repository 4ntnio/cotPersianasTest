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
    $boldStart = !empty($options['bold']) ? '<w:b/>' : '';
    $size = isset($options['size']) ? '<w:sz w:val="' . (int) $options['size'] . '"/><w:szCs w:val="' . (int) $options['size'] . '"/>' : '';
    $spacing = isset($options['spacing_after']) ? '<w:spacing w:after="' . (int) $options['spacing_after'] . '"/>' : '';
    $justify = isset($options['align']) ? '<w:jc w:val="' . escape_docx($options['align']) . '"/>' : '';

    return '<w:p><w:pPr>' . $justify . $spacing . '</w:pPr><w:r><w:rPr>' . $boldStart . $size . '</w:rPr><w:t xml:space="preserve">' . $escapedText . '</w:t></w:r></w:p>';
}

function make_docx_cell($text, $width, $header = false)
{
    $cellText = escape_docx($text);
    $runProperties = $header ? '<w:rPr><w:b/></w:rPr>' : '';

    return '<w:tc>'
        . '<w:tcPr><w:tcW w:w="' . (int) $width . '" w:type="dxa"/></w:tcPr>'
        . '<w:p><w:r>' . $runProperties . '<w:t xml:space="preserve">' . $cellText . '</w:t></w:r></w:p>'
        . '</w:tc>';
}

function make_docx_row(array $cells, $header = false)
{
    $widths = [1800, 2600, 1800, 1800, 1800];
    $xml = '<w:tr>';

    foreach ($cells as $index => $cell) {
        $xml .= make_docx_cell($cell, $widths[$index] ?? 1800, $header);
    }

    $xml .= '</w:tr>';

    return $xml;
}

function build_quote_docx_xml(array $document)
{
    $meta = $document['meta'];
    $items = $document['items'];
    $rows = [];
    $rows[] = make_docx_row(['Tipo de persiana', 'Modelo / tela', 'Color', 'Accionamiento', 'Precio por pieza'], true);

    foreach ($items as $item) {
        $rows[] = make_docx_row([
            $item['type'],
            $item['model'],
            $item['color'],
            $item['operation'],
            '$' . format_money($item['total_price']),
        ]);
    }

    $tableXml = '<w:tbl>'
        . '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
        . implode('', $rows)
        . '</w:tbl>';

    $observationText = $meta['observations'] !== '' ? $meta['observations'] : 'Sin observaciones adicionales.';

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
        . 'mc:Ignorable="w14 wp14">'
        . '<w:body>'
        . make_docx_paragraph('ZAME Blinds', ['bold' => true, 'size' => 34, 'spacing_after' => 120])
        . make_docx_paragraph('Cotización de persianas', ['bold' => true, 'size' => 28, 'spacing_after' => 220])
        . make_docx_paragraph('Folio: ' . $document['folio'], ['spacing_after' => 80])
        . make_docx_paragraph('Fecha: ' . $document['date'], ['spacing_after' => 80])
        . make_docx_paragraph('Cliente: ' . $meta['client_name'], ['spacing_after' => 80])
        . make_docx_paragraph('Teléfono / WhatsApp: ' . $meta['client_phone'], ['spacing_after' => 160])
        . $tableXml
        . make_docx_paragraph('', ['spacing_after' => 80])
        . make_docx_paragraph('Total acumulado: $' . format_money($document['total']), ['bold' => true, 'spacing_after' => 120])
        . make_docx_paragraph('Vigencia: ' . $meta['validity'], ['spacing_after' => 120])
        . make_docx_paragraph('Observaciones generales: ' . $observationText, ['spacing_after' => 120])
        . make_docx_paragraph('Cambio de precio sujeto sin previo aviso', ['bold' => true, 'spacing_after' => 120])
        . '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
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
    $documentXml = build_quote_docx_xml($document);
    $coreTitle = escape_docx('Cotización ' . $document['folio']);
    $coreSubject = escape_docx('Cotización de persianas ZAME Blinds');
    $coreAuthor = escape_docx('ZAME Blinds');

    $zip->addFromString('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        . '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        . '<Default Extension="xml" ContentType="application/xml"/>'
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
        . '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        . '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>'
        . '</w:styles>');
    $zip->addFromString('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');

    $zip->close();
}

function output_docx_download(array $items, array $meta)
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
