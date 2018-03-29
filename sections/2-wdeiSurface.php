<?php
require_once 'vendor/autoload.php';
use PhpOffice\PhpWord\Shared\Converter;

$BOP = $reportData->BOP;
//Well Data
$tableWellData = $sectionMainContent->addTable('tsPlain');
$tableWellData->addRow($rsHeight);
$tableWellData->addCell(Converter::inchtotwip(7),['gridSpan' => 2, 'valign' => 'center', 'bgColor' => '3A3B3D'])->addText("Well Data", 'fsNormal', 'psLeftTable3');
$tableWellData->addRow($rsHeight);
$tableWellData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText("Well Name", 'fsNormal', 'psLeftTable3');
$tableWellData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText($reportData->well->name, 'fsNormal', 'psLeftTable3');
$tableWellData->addRow($rsHeight);
$tableWellData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText("Well Location", 'fsNormal', 'psLeftTable3');
$tableWellData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText($reportData->well->location, 'fsNormal', 'psLeftTable3');
$tableWellData->addRow($rsHeight);
$tableWellData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText("Water Depth, Hsw (ft)", 'fsNormal', 'psLeftTable3');
$tableWellData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText($reportData->well->waterDepth, 'fsNormal', 'psLeftTable3');
$tableWellData->addRow($rsHeight);
$tableWellData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText("Maximum Anticipated Surface Pressure, MASP (psi)", 'fsNormal', 'psLeftTable3');
$tableWellData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText($reportData->well->pressure, 'fsNormal', 'psLeftTable3');
$sectionMainContent->addTextBreak(1,'fsNormal');
//Shear Ram
$tableShearRamData = $sectionMainContent->addTable('tsPlain');
$tableShearRamData->addRow($rsHeight);
$tableShearRamData->addCell(Converter::inchtotwip(7),['gridSpan' => 2, 'valign' => 'center', 'bgColor' => '3A3B3D'])->addText("Shear Ram", 'fsNormal', 'psLeftTable3');
$tableShearRamData->addRow($rsHeight);
$tableShearRamData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText("BOP OEM", 'fsNormal', 'psLeftTable3');
$tableShearRamData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText($BOP->OEM, 'fsNormal', 'psLeftTable3');
$tableShearRamData->addRow($rsHeight);
$tableShearRamData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText("Ram BOP Model", 'fsNormal', 'psLeftTable3');
$tableShearRamData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText($BOP->model, 'fsNormal', 'psLeftTable3');
$tableShearRamData->addRow($rsHeight);
$tableShearRamData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText("Supply Pressure, Ps (psi)", 'fsNormal', 'psLeftTable3');
$tableShearRamData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText($BOP->supplyPressure, 'fsNormal', 'psLeftTable3');
$tableShearRamData->addRow($rsHeight);
$tableShearRamData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText("Operator Rated Pressure, Ps (psi)", 'fsNormal', 'psLeftTable3');
$tableShearRamData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText($BOP->operatorRatedPressure, 'fsNormal', 'psLeftTable3');
$tableShearRamData->addRow($rsHeight);
$tableShearRamData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText("Shear Operator Closing Area, Ac (in2)", 'fsNormal', 'psLeftTable3');
$tableShearRamData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText($BOP->closingArea, 'fsNormal', 'psLeftTable3');
$tableShearRamData->addRow($rsHeight);
$tableShearRamData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText("Shear Operator Closing Ratio, Cr", 'fsNormal', 'psLeftTable3');
$tableShearRamData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText($BOP->closingRatio, 'fsNormal', 'psLeftTable3');
$tableShearRamData->addRow($rsHeight);
$tableShearRamData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText("Minimum Operating Pressure for Low Pressure Seal, MOPFLPS (psi)", 'fsNormal', 'psLeftTable3');
$tableShearRamData->addCell(Converter::inchtotwip(3.5),$vAlignCell)->addText($BOP->MOPFLPS, 'fsNormal', 'psLeftTable3');