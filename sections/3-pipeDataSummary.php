<?php
require_once 'vendor/autoload.php';
use PhpOffice\PhpWord\Shared\Converter;

$sectionMainContent->addTitle('Pipe Data Summary', 1);

$tablePipeDataSummary = $sectionMainContent->addTable('tsPlain');
$tablePipeDataSummary->addRow(null,$rsFirstRowConfig);  //I guess I can't configure a row.
$tablePipeDataSummary->addCell(Converter::inchtotwip(.58),$vAlignCell)->addText("No.", 'fsNormal', 'psCenterTable3');
$tablePipeDataSummary->addCell(Converter::inchtotwip(.65),$vAlignCell)->addText("Grade", 'fsNormal', 'psCenterTable3');
$tablePipeDataSummary->addCell(Converter::inchtotwip(.54),$vAlignCell)->addText("OD (in)", 'fsNormal', 'psCenterTable3');
$tablePipeDataSummary->addCell(Converter::inchtotwip(.83),$vAlignCell)->addText("Wall Thickness (in)", 'fsNormal', 'psCenterTable3');
$tablePipeDataSummary->addCell(Converter::inchtotwip(.77),$vAlignCell)->addText("Strength", 'fsNormal', 'psCenterTable3');
$tablePipeDataSummary->addCell(Converter::inchtotwip(.72),$vAlignCell)->addText("Nominal Weight (ppf)", 'fsNormal', 'psCenterTable3');
$tablePipeDataSummary->addCell(Converter::inchtotwip(.89),$vAlignCell)->addText("Max. Calculated Shear/Seal Pressure (psi)", 'fsNormal', 'psCenterTable3');
$tablePipeDataSummary->addCell(Converter::inchtotwip(.75),$vAlignCell)->addText("Actual Shear Pressure (psi)", 'fsNormal', 'psCenterTable3');
$tablePipeDataSummary->addCell(Converter::inchtotwip(.9),$vAlignCell)->addText("Adjusted Actual Shear Pressure (psi)", 'fsNormal', 'psCenterTable3');
$tablePipeDataSummary->addCell(Converter::inchtotwip(.82),$vAlignCell)->addText("Verification Status", 'fsNormal', 'psCenterTable3');