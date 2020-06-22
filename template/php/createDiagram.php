<?php
/**
 * Creates a excel diagram out of inputted data
 *
 *
 * PHP version 7.3.6
 *
 *
 * @package projectEval
 * @author Original Author <justin.thiede@visions.ch>
 * @copyright visions.ch GmbH
 * @license http://creativecommons.org/licenses/by-nc-sa/3.0/
 */

require_once 'libs/phpoffice_phpspreadsheet/vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Chart\DataSeries;
use PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;
use PhpOffice\PhpSpreadsheet\Chart\PlotArea;
use PhpOffice\PhpSpreadsheet\Chart\Layout;
use PhpOffice\PhpSpreadsheet\Chart\Legend;
use PhpOffice\PhpSpreadsheet\Chart\Title;
use PhpOffice\PhpSpreadsheet\Chart\Chart;
use PhpOffice\PhpSpreadsheet\Chart\GridLines;
use PhpOffice\PhpSpreadsheet\Helper\Html;

foreach ($_POST as $key => $value) {
    $$key = $value;
}

$htmlHelper     = new Html;
$letters        = range('A','Z');
$header         = [''];
$headerFirst    = 'B';
$headerQuantity = 0; // Amount of headers

// Create header array and set its starting and ending cells
foreach ($ticketId as $index => $id) {
    $text     = 'Description: <b>' . $description[$index] . '</b><br>Ticket: <b>#' . $id . '</b><br>Assigned to: <b>' . $assignedTo[$index] . '</b>';
    $richText = $htmlHelper->toRichTextObject($text);

    array_push($header, $richText);
    $headerQuantity++;
}

// Add last column for total overview
array_push($header, 'Total');
array_push($estimatedEffort, array_sum($estimatedEffort));
array_push($timeSpent, array_sum($timeSpent));

$headerQuantity++;
$cellQuantity   = count($estimatedEffort) - 1;
$columnLast     = $letters[$cellQuantity];

$spreadsheet = new Spreadsheet();
$worksheet   = $spreadsheet->getActiveSheet();
$worksheet->fromArray(
    [
        $header,
        $estimatedEffort,
        $timeSpent
    ]
);

$dataSeriesLabels = [
    new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$A$2', null, 1),
    new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$A$3', null, 1),
];

$xAxisTickValues = [
    new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$' . $headerFirst . '$1:$' . $letters[$headerQuantity] . '$1', null, $headerQuantity),
];

$dataSeriesValues = [
    new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_NUMBER, 'Worksheet!$B$2:$' . $columnLast . '$2', null, $cellQuantity),
    new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_NUMBER, 'Worksheet!$B$3:$' . $columnLast . '$3', null, $cellQuantity),
];

$series = new DataSeries(
    DataSeries::TYPE_BARCHART,
    DataSeries::GROUPING_STANDARD,
    range(0, count($dataSeriesValues) - 1),
    $dataSeriesLabels,
    $xAxisTickValues,
    $dataSeriesValues,
);

$series->setPlotDirection(DataSeries::DIRECTION_COLUMN);
$layout   = new Layout();
$layout->setShowVal(true); // Show data value in chart
$plotArea = new PlotArea($layout, [$series]); // Set the series in the plot area
$legend   = new Legend(Legend::POSITION_BOTTOM, null, false); // Set the chart legend
$title    = new Title($project);
$chart    = new Chart(
    'chart1',
    $title,
    $legend,
    $plotArea,
    true,
    0,
    null,
    null
); // Create the chart

// Set the position where the chart should appear in the worksheet
$chart->setTopLeftPosition('D9');
$chart->setBottomRightPosition('K22');

$worksheet->addChart($chart); // Add the chart to the worksheet
$filename = 'project_eval';
$writer   = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->setIncludeCharts(true);

header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="'. $filename .'.xlsx"');
header('Cache-Control: max-age=0');

$writer->save('php://output'); // download file
