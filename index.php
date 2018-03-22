<?php
require_once 'vendor/autoload.php';
use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpWord\Style\Paper;

// Creating the new document...
$phpWord = new \PhpOffice\PhpWord\PhpWord();
$phpWord->getCompatibility()->setOoxmlVersion(15);

$properties = $phpWord->getDocInfo();
$properties->setCreator('David Hanks');
$properties->setCompany('OTC Solutions');
$properties->setTitle('Shear Verification Package');
$properties->setDescription('My description');
$properties->setLastModifiedBy('My name');
$properties->setCreated(mktime(0, 0, 0, 3, 12, 2014));

//
//Styles
//

//Fonts
$fontStyleNormalname = 'fsNormal';
$phpWord->addFontStyle($fontStyleNormalname, array(
	'name' => 'Arial',
	'size' => 10
));

$phpWord->addFontStyle('fsTitle', array(
	'name' => 'Arial',
	'size' => 20,
	'bold' => true,
	
));

$phpWord->addFontStyle('fsSubTitle', array(
	'name' => 'Arial',
	'size' => 16,
	'bold' => true,
	'spaceAfter' => 0,
	'spaceBefore' => 0
));

//Paragraphs
$styleParaNormal = array(
	'alignment' => 'justified'
);

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

$borderSize = 5;
$styleSectionTitle = array_merge($styleSectionDefault, ['borderBottomSize' => $borderSize,'borderLeftSize' => $borderSize,'borderRightSize' => $borderSize,'borderTopSize' => $borderSize,]);

$sectionTitlePg = $phpWord->addSection($styleSectionTitle);

//Add the coverpage border


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
			'width' => $width,//7.3"
			'height' => $height//5.92"
        )
);

//Add the coverpage Title, Rig, Client, Doc#, Rev, Release Date
$sectionTitlePg->addText('Shear Verification Package', 'fsTitle', $styleParaTitle);
$sectionTitlePg->addText('30 CFR §250.732 (a),(b)', 'fsSubTitle', $styleParaSubTitle);
$sectionTitlePg->addTextBreak(1);
$sectionTitlePg->addText('[Client Legal Name]', 'fsTitle', $styleParaTitle);
$sectionTitlePg->addText('[Rig]', 'fsSubTitle', $styleParaSubTitle);
$sectionTitlePg->addText('[Well Name]', 'fsSubTitle', $styleParaSubTitle);
$sectionTitlePg->addTextBreak(1);
$sectionTitlePg->addText('[Document No]', 'fsSubTitle', $styleParaSubTitle);
$sectionTitlePg->addText('[REV]', 'fsSubTitle', $styleParaSubTitle);
$sectionTitlePg->addText('March 22, 2018', 'fsSubTitle', $styleParaSubTitle);



//add the Coverpage OTCSolution image
$sectionTitlePg->addImage(
		'img/OTCsolutions_cover.jpg',
		array(
            'height' => Converter::inchToPixel(.977),
			'positioning' => \PhpOffice\PhpWord\Style\Image::POSITION_ABSOLUTE,
            'marginTop'          => Converter::inchToTwip(-8),
			'marginLeft' => Converter::inchToPixel(-.6),
            'wrappingStyle' => 'tight'
        )
);


// Saving the document as OOXML file...
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save('output/helloWorld.docx');

?>