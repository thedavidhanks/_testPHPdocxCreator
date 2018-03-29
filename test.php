<?php
require_once 'vendor/autoload.php';

function comparePipeNo($a,$b){
	if ($a->pipeNo == $b->pipeNo) {
        return 0;
    }
    return ($a->pipeNo < $b->pipeNo) ? -1 : 1;
}

function compareArrgh($a,$b){
	if ($a == $b) {
        return 0;
    }
    return ($a < $b) ? -1 : 1;
}
// Creating the new document...
//$phpWord = new \PhpOffice\PhpWord\PhpWord();

//The function will receive a JSON as an argument
$JSONdata = file_get_contents('testing/setup1.json');
$reportData = json_decode($JSONdata);

$pipesObj = $reportData->tubulars;
$newPipeArray = get_object_vars($pipesObj);
usort($newPipeArray, "comparePipeNo");
foreach($newPipeArray as $data){
	echo $data->pipeNo.',';
}

?>