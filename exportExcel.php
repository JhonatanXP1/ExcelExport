<?php
// Manejar solicitud OPTIONS (preflight) primero
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    header('Access-Control-Allow-Origin: *');
    header('Access-Control-Allow-Methods: POST, OPTIONS');
    header('Access-Control-Allow-Headers: Content-Type, Authorization, X-Requested-With');
    header('Access-Control-Max-Age: 86400'); // 24 horas
    header('HTTP/1.1 200 OK');
    exit();
}

// Para solicitudes que no son OPTIONS, aplicar verificación normal
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: POST');
header('Access-Control-Allow-Headers: Content-Type, Authorization, X-Requested-With');

// Verificación de sesión (solo para POST)
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (!headers_sent() && session_status() === PHP_SESSION_NONE) session_start();
} else {
    header('HTTP/1.1 401 Unauthorized');
    header('WWW-Authenticate: Basic realm="Acceso restringido"');
    echo 'Acceso no autorizado';
    exit;
}
#[NoReturn] // Si algo falla se notificara al front.
function ReportServer500($error, $code = ""): void
{
    header('HTTP/1.1 500 Internal Server Error');
    echo "error en el servidor: $code - $error";
    exit;
}

error_reporting(E_ALL);
ini_set('display_errors', 1);
error_reporting(E_ERROR);
ini_set('display_errors', FALSE);
ini_set('log_errors', TRUE);
set_time_limit(0);
ignore_user_abort(false);
ini_set('memory_limit', '-1');

$input = file_get_contents('php://input');
$datos = json_decode($input, true);

require_once __DIR__ . '/vendor/autoload.php';
require_once $_SERVER['DOCUMENT_ROOT'] . '/reportes/utf8EncodeDecode.php';

if ($datos['encode'] && array_key_exists('encode', $datos)) {
    $datos = utf8ize_decode($datos);
}

use JetBrains\PhpStorm\NoReturn;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Exception as PhpSpreadsheetException;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

$headerStyle = [
    'font' => [
        'bold' => true,
        'size' => 13,
        'color' => ['argb' => 'FFFFFFFF']
    ],
    'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'startColor' => ['argb' => 'FF4472C4']
    ],
    'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
        'wrapText' => true
    ]
];

$bodyStyle = [
    'font' => [
        'size' => 12,
        'color' => [],
    ],
    'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
        'wrapText' => true
    ],
    'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'startColor' => []
    ]
];

/**
 * @throws Exception
 */

function currice($positionX, $positionY, $celda): void
{
    global $sheetName, $sheet, $indiceX;
    if (!is_numeric($celda)) {
        throw new Exception("\nthe sheet:$sheetName\nPosition:[" . ($positionY - 1) . "][$indiceX]\nvalue:$celda[0] is not numeric");
    }

    $sheet->setCellValueExplicit($positionX . $positionY + 1, $celda, DataType::TYPE_NUMERIC);
    $sheet->getStyle($positionX . $positionY + 1)
        ->getNumberFormat()
        ->setFormatCode('"$ "#,##0.00');
}

/**
 * @throws Exception
 */
function CreateExcel(Worksheet $sheet, $dataArraySheet): void
{
    global $titulo, $headerStyle, $bodyStyle, $indice, $fillter;
    //Si en algún momento un encabezado tiene un merge se hace un conteo extra por las columnas mergeadas.
    $mergeXheader = 0;
    $columnCount = $indice ? 1 : 0;
    $colorRegitrosTable = false;
    $extrPropier = array_filter($dataArraySheet[0], 'is_array');
    if (!empty($extrPropier)) {
        foreach ($extrPropier as $existMerge) {
            if (array_key_exists("mergeX", $existMerge))
                $mergeXheader += $existMerge["mergeX"];
        }
    }

    if ($indice) {
        $mergeXheader++;
        $col = Coordinate::stringFromColumnIndex($columnCount);
        $sheet->setCellValue($col . "2", "No.");
        $sheet->getStyle($col . "2")->applyFromArray($headerStyle);
    }
    /******************************************** START ENCABEZADO GENERAL *********************************************************/
    $finalHeaderGlobal = Coordinate::stringFromColumnIndex(count($dataArraySheet[0]) + $mergeXheader);
    $sheet->mergeCells('A1:' . $finalHeaderGlobal . '1');
    $sheet->setCellValue('A1', $titulo);//Definicion de Encabezado Global;

    $sheet->getStyle('A1')
        ->applyFromArray([
            'font' => [
                'bold' => true,
                'size' => 13,
                'color' => ['argb' => 'FFFFFFFF'],
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['argb' => 'FF314D85'],
            ],
        ]);
    /******************************************* END ENCABEZADO GENERAL ***********************************************************/

    foreach ($dataArraySheet[0] as $encabezado) { //Definir de tabla Encabezados;
        $columnCount = $columnCount + 1;
        $col = Coordinate::stringFromColumnIndex($columnCount);
        if (is_scalar($encabezado)) {
            $sheet->setCellValue($col . "2", "$encabezado");
            $sheet->getStyle($col . "2")->applyFromArray($headerStyle);
        }
        if (gettype($encabezado) == "array") {
            if (array_key_exists('mergeX', $encabezado)) { // Verificas que si una celda tiene una propiedad que defina un merge a lo vertical.
                $fillter = false;
                $colFinal = Coordinate::stringFromColumnIndex($columnCount + $encabezado['mergeX']);
                $sheet->mergeCells($col . '2:' . $colFinal . '2');
                $columnCount += $encabezado['mergeX'];
            }
            $sheet->setCellValue($col . "2", $encabezado[0]);
            $sheet->getStyle($col . "2")->applyFromArray($headerStyle);
        }
    }

    $sheet->freezePane('A3');

    array_shift($dataArraySheet);
    $indexRow = 0;
    foreach ($dataArraySheet as $indexY => $rowValues) {
        $indexXtemp = $indice ? 1 : 0;
        $bodyStyle['font']['color'] = $colorRegitrosTable ? ['argb' => 'FF303030'] : ['argb' => 'FF181818'];
        $bodyStyle['fill']['startColor'] = $colorRegitrosTable ? ['argb' => 'FFC9D6EE'] : ['argb' => 'FFFAFAFA'];
        $tempIndexRow = "";
        if ($indice) {
            $indexRow++;
            $indexXIndice = Coordinate::stringFromColumnIndex($indexXtemp);
            $sheet->setCellValue($indexXIndice . $indexY + 3, $indexRow);
            $sheet->getStyle($indexXIndice . $indexY + 3)->applyFromArray($bodyStyle);
            $tempIndexRow = $indexXIndice . $indexY + 3;
        }
        $indexY += 2;

        foreach ($rowValues as $indiceX => $celda) {
            $indexXtemp += 1;
            $columIndex = Coordinate::stringFromColumnIndex($indexXtemp);
            if (is_scalar($celda)) {
                $sheet->setCellValue($columIndex . $indexY + 1, "$celda");
                $sheet->getStyle($columIndex . $indexY + 1)->applyFromArray($bodyStyle);
            }

            if (gettype($celda) == "array") {
                $temBodyStyle = $bodyStyle;
                $curriceActive = array_key_exists('currency', $celda);
                $mergeYActive = array_key_exists('mergeY', $celda);
                if (array_key_exists('mergeX', $celda)) {
                    $fillter = false;
                    $indexXfinal = Coordinate::stringFromColumnIndex($indexXtemp + $celda['mergeX']);
                    $sheet->mergeCells($columIndex . ($indexY + 1) . ':' . $indexXfinal . ($indexY + 1));
                    $indexXtemp += $celda['mergeX'];
                }
                if (array_key_exists('mergeY', $celda)) {
                    $fillter = false;
                    $sheet->mergeCells($columIndex . ($indexY + 1 - $celda['mergeY']) . ":" . $columIndex . ($indexY + 1));
                }

                if ($indice && isset($celda['indice']) && !$celda['indice'] && $tempIndexRow) { // Si encuentra la bandera [indice], elimina directamente el index de la fila y reestablece el contador.
                    $sheet->setCellValue($tempIndexRow, "");
                    $tempIndexRow = "";
                    $indexRow--;
                }

                if (($celda['strong'] ?? false) === true)
                    $temBodyStyle['font']['bold'] = true;

                if (isset($celda['color']))
                    $temBodyStyle['font']['color'] = ['argb' => "FF{$celda['color']}"];

                if (isset($celda['fillColor']))
                    $temBodyStyle['fill']['startColor'] = ['argb' => "FF{$celda['fillColor']}"];

                if (isset($celda['align']))
                    $temBodyStyle['alignment']['horizontal'] = ($celda['align'] == 'left') ? Alignment::HORIZONTAL_LEFT : Alignment::HORIZONTAL_RIGHT;


                if (!$mergeYActive && !$curriceActive) {
                    $sheet->setCellValue($columIndex . $indexY + 1, $celda[0]);
                    $sheet->getStyle($columIndex . $indexY + 1)->applyFromArray($temBodyStyle);
                } else {
                    if ($curriceActive) {// Aplica formato de moneda (MXN) manteniendo el valor como numérico para permitir cálculos en Excel.
                        currice($columIndex, $indexY, $celda[0]);
                        $sheet->getStyle($columIndex . $indexY + 1)->applyFromArray($temBodyStyle);
                    }
                    if ($mergeYActive) {
                        if ($curriceActive) {
                            currice($columIndex, ($indexY - $celda['mergeY']), $celda[0]);
                        } else {
                            $sheet->setCellValue($columIndex . ($indexY + 1 - $celda['mergeY']), $celda[0]);
                        }
                        $sheet->getStyle($columIndex . ($indexY + 1 - $celda['mergeY']))->applyFromArray($temBodyStyle);
                    }
                }
            }
        }
        $colorRegitrosTable = !$colorRegitrosTable;
    }

    if ($fillter) {
        $positionY = count($dataArraySheet) + 2;
        $sheet->setAutoFilter("A2:$finalHeaderGlobal$positionY");
    }
    $highestColumn = $sheet->getHighestColumn();
    foreach (range('A', $highestColumn) as $col) {
        $sheet->getColumnDimension($col)->setAutoSize(true);
    }
}

try {
    $nameFile = $datos["nameFile"] ?? "";
    $titulo = $datos['title'] ?? "";
    $indice = $datos['indice'] ?? false;
    $fillter = $datos['fillter'] ?? false;
    $spreadsheet = new Spreadsheet();
    if (!empty($datos)) {
        if (gettype($datos) == "array") {
            $spreadsheet->getProperties()
                ->setCreator("Ing. Jhonatan")
                ->setLastModifiedBy("Ing. Jhonatan")
                ->setSubject("Reportes");

            if ($datos['data'] != null) {
                if (is_numeric(array_key_first($datos['data'])) && count($datos['data']) == 1) { // Es porque es una simple tabla.
                    $sheet = $spreadsheet->getActiveSheet();
                    $sheet->setTitle(0);
                    CreateExcel($sheet, $datos['data'][0]);
                } else { // Es porque serán más de 1 y posiblemente tenga sheet personalizados;
                    $indiceSheet = 0;
                    foreach ($datos['data'] as $sheetName => $dataExcel) {
                        $nameSheet = substr(str_replace(" ", "", $sheetName), 0, 30);
                        $nameSheet = iconv('UTF-8', 'ASCII//TRANSLIT', $nameSheet);
                        $nameSheet = preg_replace('/[\\\\\/:*?\[\]]/', '_', $nameSheet);
                        if ($indiceSheet) {
                            $sheet = new Worksheet($spreadsheet, $nameSheet);
                            $spreadsheet->addSheet($sheet, $indiceSheet);
                        } else {
                            $sheet = $spreadsheet->getActiveSheet();
                            $sheet->setTitle($nameSheet);
                        }
                        CreateExcel($sheet, $dataExcel);
                        $indiceSheet++;
                    }
                }
            } else {
                ReportServer500("El Data del Array no posee datos");
            }
        }
    } else {
        ReportServer500("No se encontro el array");
    }

    $writer = new Xlsx($spreadsheet);
    $nameFile = utf8_decode($nameFile);
    $nameFile = $nameFile . date('Y-m-d') . '.xlsx';
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header("Content-Disposition: attachment; filename=\"$nameFile\"");
    header("Expires: 0");
    header("Cache-Control: no-store, no-cache, must-revalidate, max-age=0");
    header("Pragma: no-cache");
    $writer->save('php://output');
} catch (PhpSpreadsheetException $errorPhpSheet) {
    ReportServer500($errorPhpSheet->getMessage(), $errorPhpSheet->getCode());
} catch (Exception $e) {
    reportServer500($e->getMessage(), $e->getCode());
}