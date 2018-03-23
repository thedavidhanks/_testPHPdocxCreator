<?php
require_once 'vendor/autoload.php';
use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpWord\Style\Paper;


function reportDocxCreator($data) {
    
    $reportData = json_decode($data);
    $revisionNumber="Rev ".$reportData->docRev;
    
    // Creating the new document...
    $phpWord = new \PhpOffice\PhpWord\PhpWord();
    $phpWord->getCompatibility()->setOoxmlVersion(15);

    $properties = $phpWord->getDocInfo();
    $properties->setCreator($reportData->author);
    $properties->setCompany($reportData->company);
    $properties->setTitle($reportData->reportType);
    $properties->setDescription($reportData->docDescription);
    $properties->setLastModifiedBy($reportData->lastModifier);
    $properties->setCreated(mktime(0, 0, 0, 3, 12, 2014));

    //
    //Styles
    //

    //Fonts
    $phpWord->addFontStyle('fsNormal', array(
            'name' => 'Arial',
            'size' => 10
    ));

    $phpWord->addFontStyle('fsTitle', array(
            'name' => 'Arial',
            'size' => 20,
            'bold' => true,

    ));
    
    $phpWord->addFontStyle('fsHeader', array(
        'name' => 'Arial',
        'size' => 8
    ));

    $phpWord->addFontStyle('fsSubTitle', array(
            'name' => 'Arial',
            'size' => 16,
            'bold' => true,
            'spaceAfter' => 0,
            'spaceBefore' => 0
    ));
    $phpWord->addFontStyle('fsRevBlock', array(
        'name' => 'Arial',
        'size' => 14,
        'bold' => true,
        'color' => '8B2027'
    ));

    //Paragraphs
    $phpWord->addParagraphStyle('Normal', array(
            'alignment' => 'justified'
    ));

    $styleParaTitle = array(
            'alignment' => 'right',
            'spaceAfter' => 2,
            'spaceBefore' => 2
    );

    $styleParaSubTitle = array(
            'alignment' => 'right',
            'spaceAfter' => 0,
            'spaceBefore' => 0
    );
    $styleParaRevs = array(
            'basedOn' => 'Normal',
            'alignment' => 'left'
    );

    //Sections
    $styleSectionDefault = array(
            'paperSize' => 'Letter',
            'marginTop' => Converter::inchToTwip(.5),
            'marginLeft' => Converter::inchToTwip(.5),
            'marginRight' => Converter::inchToTwip(.5),
            'marginBottom' => Converter::inchToTwip(.3),
            'headerHeight' => Converter::inchToTwip(.3),
            'footerHeight' => Converter::inchToTwip(.4)
    );

    $styleSectionTitle = array_merge($styleSectionDefault, [
        'borderSize' => 40,
        'borderColor' => '9E3234'
        ]
    );

    /*
     * BEGIN DOC CREATION
     */
    
    $sectionTitlePg = $phpWord->addSection($styleSectionTitle);

    //Add the coverpage image
    $width = 525;
    $height = 1302/1605*$width;
    $sectionTitlePg->addImage(
        'img/drivenByRealExp.jpg',
        array(
            'positioning'      => \PhpOffice\PhpWord\Style\Image::POSITION_ABSOLUTE,
            'posHorizontal'    => \PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_CENTER,
            'posHorizontalRel' => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE_TO_PAGE,
            'marginTop'          => 0.75,
            'wrappingStyle'      => 'square',
            'width' => $width,//7.3" or 1605 pixels
            'height' => $height//5.92" or 1302 pixels
        )
    );

    //Add the coverpage Title, Rig, Client, Doc#, Rev, Release Date
    $sectionTitlePg->addText($reportData->reportType, 'fsTitle', $styleParaTitle);
    $sectionTitlePg->addText($reportData->reportSubType, 'fsSubTitle', $styleParaSubTitle);
    $sectionTitlePg->addTextBreak(1);
    $sectionTitlePg->addText('[Client Legal Name]', 'fsTitle', $styleParaTitle);
    $sectionTitlePg->addText('[Rig]', 'fsSubTitle', $styleParaSubTitle);
    $sectionTitlePg->addText('[Well Name]', 'fsSubTitle', $styleParaSubTitle);
    $sectionTitlePg->addTextBreak(1);
    $sectionTitlePg->addText($reportData->docNumber, 'fsSubTitle', $styleParaSubTitle);
    $sectionTitlePg->addText($revisionNumber, 'fsSubTitle', $styleParaSubTitle);
    $sectionTitlePg->addText('March 22, 2018', 'fsSubTitle', $styleParaSubTitle);



    //add the Coverpage OTCSolution image
    $sectionTitlePg->addImage(
        'img/OTCsolutions_cover.jpg',
        array(
            'height' => Converter::inchToPixel(.977),
            'wrappingStyle'      => 'square',
            'positioning' => \PhpOffice\PhpWord\Style\Image::POSITION_ABSOLUTE,
            'posVertical'      => \PhpOffice\PhpWord\Style\Image::POSITION_VERTICAL_BOTTOM,
            'posHorizontal' =>\PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_LEFT,
            'marginLeft'       => .5,
            'marginBottom'        => 0
        )
    );

    //
    //REVISION PAGE
    $sectionRevPage=$phpWord->addSection($styleSectionDefault);
    $header=$sectionRevPage->addHeader();
    $header->addText($reportData->reportType, 'fsHeader', $styleParaSubTitle);
    $header->addText($reportData->docNumber, 'fsHeader', $styleParaSubTitle);
    $header->addText($revisionNumber, 'fsHeader', $styleParaSubTitle);
        //add the Coverpage OTCSolution image
    $header->addImage(
        'img/OTCsolutions_header.jpg',
        array(
            'height' => Converter::inchToPixel(.5625),
            'wrappingStyle'      => 'square',
            'positioning' => \PhpOffice\PhpWord\Style\Image::POSITION_ABSOLUTE,
            'posVerticalRel'   => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE_TO_PAGE,
            'posHorizontal' =>\PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_LEFT,
            'marginLeft'       => .5,
            'marginTop' => .25
        )
    );
    
    $sectionRevPage->addTextBreak(1, 'fsNormal');
    $sectionRevPage->addText('REVISION HISTORY', 'fsRevBlock', $styleParaRevs);
    //add table
    $sectionRevPage->addText('CHANGE DESCRIPTION', 'fsRevBlock', $styleParaRevs);
    //add change description table
    $tableChangeDesc = $sectionRevPage->addTable();
    for ($r = 1; $r <= 2; $r++) {
    $tableChangeDesc->addRow();
    for ($c = 1; $c <= 2; $c++) {
        $table->addCell(1750)->addText("Row {$r}, Cell {$c}");
    }
}
    
    
    // Saving the document as OOXML file...
    $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
    $objWriter->save('output/helloWorld.docx');
}

//The function will receive a JSON as an argument
$createDate = mktime(0, 0, 0, 3, 12, 2014);
$jsonArgument = <<<EOT
    {     
        "author": "David Hanks",
        "company" : "OTC Solutions",
        "docDescription" : "",
        "lastModifier" : "",
        "reportType" : "Shear Verification Package",
        "reportSubType" : "30 CFR §250.732 (a),(b)",
        "clientFullLegal" : "",
        "clientShortend" : "OES",
        "rig" : "",
        "wellName" : "",
        "docNumber": "SBS1_COX_180307_SVP",
        "docRev": 0,
        "revisions" : {
            "00" : {
                "date" : "2016-01-01",
                "descShort" : "Initial Release",
                "descLong" : "Inital Release",
                "preapredBy" : "David Hanks",
                "checkedBy" : "Steve Ronan",
                "approvedBy" : "David Hanks"
            },
            "01" : {
                    "date" : "2016-02-20",
                    "descShort" : "Fixed it",
                    "descLong" : "Something was really, really wrong.  It had to go.  It's all better now",
                    "preapredBy" : "David Hanks",
                    "checkedBy" : "Steve Ronan",
                    "approvedBy" : "David Hanks"
                }
        }
    }
EOT;
//echo $jsonArgument;
//$reportData = json_decode($jsonArgument);

reportDocxCreator($jsonArgument);
?>