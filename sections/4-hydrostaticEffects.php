<?php
require_once 'vendor/autoload.php';
use PhpOffice\PhpWord\Shared\Converter;

function addTermCell($table, $term = '', $sub = ''){
    global $vAlignCell;
    
    $trTerm=$table->addCell(null, $vAlignCell)->addTextRun('psCenterTight');
    $trTerm->addText($term,'fsSize9');
    $trTerm->addText($sub, 'fsNormalSub');
    return $table;
}

function addSqInchCell($table){
    global $vAlignCell;
    
    $trIN2=$table->addCell(null,$vAlignCell)->addTextRun('psCenterTight');
    $trIN2->addText('in','fsSize9');
    $trIN2->addText('2','fsNormalSup');
    
    return $table;
}

//Cell style
$csSpanRow = ['gridSpan' => 5, 'valign' => 'center', 'bgColor' => 'BFBFBF'];

$tableHSeffects = $sectionMainContent->addTable('tsPlain');
$tableHSeffects->addRow($rsHeight);
$tableHSeffects->addCell(Converter::inchtotwip(2.55),$csVbottomGrey)->addText("Definition", 'fsNormal', 'psCenterTable3');
$tableHSeffects->addCell(Converter::inchtotwip(.69),$csVbottomGrey)->addText("Term", 'fsNormal', 'psCenterTable3');
$tableHSeffects->addCell(Converter::inchtotwip(2.58),$csVbottomGrey)->addText("Equation", 'fsNormal', 'psCenterTable3');
$tableHSeffects->addCell(Converter::inchtotwip(.71),$csVbottomGrey)->addText("Result", 'fsNormal', 'psCenterTable3');
$tableHSeffects->addCell(Converter::inchtotwip(.46),$csVbottomGrey)->addText("Units", 'fsNormal', 'psCenterTable3');
$tableHSeffects->addRow($rsHeight);
$tableHSeffects->addCell(null,$csSpanRow)->addText("BOP Calculations", 'fsNormalBold', 'psLeftTight');

//Closing Area
$tableHSeffects->addRow($rsHeight);
$tableHSeffects->addCell(null, $vAlignCell)->addText("Closing Area",'fsSize9', 'psLeftTight');
addTermCell($tableHSeffects, 'A', 'C');
$tableHSeffects->addCell(null, $csGrey);
$tableHSeffects->addCell(null,$vAlignCell)->addText($reportData->BOP->closingArea, 'fsSize9', 'psCenterTight');
addSqInchCell($tableHSeffects);

//Closing Ratio
$tableHSeffects->addRow($rsHeight);
$tableHSeffects->addCell(null, $vAlignCell)->addText("Closing Ratio",'fsSize9', 'psLeftTight');
addTermCell($tableHSeffects, 'C', 'R');
$tableHSeffects->addCell(null, $csGrey);
$tableHSeffects->addCell(null,$vAlignCell)->addText($reportData->BOP->closingRatio, 'fsSize9', 'psCenterTight');
$tableHSeffects->addCell(null, $csGrey);

//BOP Calculations for Subsea BOP
if($reportData->rig->location == 'subsea'){
    //Tailrod Area
    $tableHSeffects->addRow($rsHeight);
    $tableHSeffects->addCell(null, $vAlignCell)->addText("Tail Rod Area",'fsSize9', 'psLeftTight');
    addTermCell($tableHSeffects, 'A', 'T');
    $tableHSeffects->addCell(null, $csGrey);
    $tableHSeffects->addCell(null,$vAlignCell)->addText($reportData->BOP->trArea, 'fsSize9', 'psCenterTight');
    addSqInchCell($tableHSeffects);
    
    //Ramshaft Area
    $tableHSeffects->addRow($rsHeight);
    $tableHSeffects->addCell(null, $vAlignCell)->addText("Ramshaft Area",'fsSize9', 'psLeftTight');
    addTermCell($tableHSeffects, 'A', 'R');
    $equationRamShaft = $tableHSeffects->addCell(null, $vAlignCell)->addTextRun('psLeftTight');
    $equationRamShaft->addText('A');
    $equationRamShaft->addText('C', 'fsNormalSub');
    $equationRamShaft->addText(' / C');
    $equationRamShaft->addText('R', 'fsNormalSub');
    $tableHSeffects->addCell(null,$vAlignCell)->addText($reportData->BOP->trArea, 'fsSize9', 'psCenterTight');
    addSqInchCell($tableHSeffects);
    //Opening Area
    $tableHSeffects->addRow($rsHeight);
    $tableHSeffects->addCell(null, $vAlignCell)->addText("Opening Area",'fsSize9', 'psLeftTight');
    addTermCell($tableHSeffects, 'A', 'O');
    $equationOpeningArea = $tableHSeffects->addCell(null, $vAlignCell)->addTextRun('psLeftTight');
    $equationOpeningArea->addText('(A');
    $equationOpeningArea->addText('C', 'fsNormalSub');
    $equationOpeningArea->addText(' + A');
    $equationOpeningArea->addText('T', 'fsNormalSub');
    $equationOpeningArea->addText(') - A');
    $equationOpeningArea->addText('R', 'fsNormalSub');
    $tableHSeffects->addCell(null,$vAlignCell)->addText($reportData->BOP->openingArea, 'fsSize9', 'psCenterTight');
    addSqInchCell($tableHSeffects);
}

$tableHSeffects->addRow($rsHeight);
$tableHSeffects->addCell(null,$csSpanRow)->addText("Wellbore Pressure", 'fsNormalBold', 'psLeftTight');