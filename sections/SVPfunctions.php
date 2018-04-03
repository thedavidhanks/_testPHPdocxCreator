<?php
require_once 'vendor/autoload.php';
use PhpOffice\PhpWord\Shared\Converter;

function generateHeader($header, $reportData, $revisionNumber){
    $headerTable = $header->addTable('tsFooter');
    $headerTable->addRow();
    $headerTable->addCell(null)->addImage(
        'img/OTCsolutions_header.jpg',
        [
        'height' => Converter::inchToPixel(.5625),
        'wrappingStyle'=> 'square',
        'positioning' => 'relative',
        'posHorizontal' =>\PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_LEFT
        ]
    );
    $headerRtCell = $headerTable->addCell(Converter::inchToTwip(4.34));
    $headerRtCell->addText($reportData->reportType, 'fsHeader', 'psRightTight');
    $headerRtCell->addText($reportData->docNumber, 'fsHeader', 'psRightTight');
    $headerRtCell->addText($revisionNumber, 'fsHeader', 'psRightTight');
    return $header;
}

