<?php 
require_once 'bootstrap.php';
$phpWord = new \PhpOffice\PhpWord\PhpWord();

/**add style section */
$styleSection = array(
    'orientation' => 'landscape'
);

$section = $phpWord->addSection();

$fontStyle = new \PhpOffice\PhpWord\Style\Font();
$fontStyle->setBold(true);
$fontStyle->setName('Tahoma');
$fontStyle->setSize(16);
$phpWord->addTitleStyle(0, $fontStyle);
$section->addTitle('Hello World');

$section->addText(
    'Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry\'s standard dummy 
    text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. '
);
for($i=1;$i<=5;$i++){
$section->addListItem('list item ' . $i);
$section->addListItem('list item ' . $i,1);
$section->addListItem('list item ' . $i,2);
}

$phpWord->addNumberingStyle(
    'multilevel',
    array(
        'type' => 'multilevel',
        'levels' => array(
            array('format' => 'decimal', 'text' => '%1.', 'left' => 360, 'hanging' => 360, 'tabPos' => 360),
            array('format' => 'upperLetter', 'text' => '%2.', 'left' => 720, 'hanging' => 360, 'tabPos' => 720),
        )
    )
);
$section->addListItem('List Item I', 0, null, 'multilevel');
$section->addListItem('List Item I.a', 1, null, 'multilevel');
$section->addListItem('List Item I.b', 1, null, 'multilevel');
$section->addListItem('List Item II', 0, null, 'multilevel');

/** set properties file */
$properties = $phpWord->getDocInfo();
$properties->setCreator('username');
$properties->setCompany('company');
$properties->setTitle('Dokumen Microsoft Word');

// unlink('helloWorld.docx');
// Saving the document as OOXML file...
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save('helloWorld.docx');

