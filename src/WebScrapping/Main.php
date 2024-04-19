<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * Runner for the Webscrapping exercice.
 */
class Main {

  /**
   * Main runner, instantiates a Scrapper and runs.
   */
  public static function run(): void {
    libxml_use_internal_errors(use_errors: true);

    $html = file_get_contents('assets/origin.html');
    //echo $html;
    echo "documento html capturado \n";
    $dom = new DOMDocument('1.0', 'utf-8');
    $dom->loadHTML($html);
    echo "carregando documento html \n";
    
     // Write your logic to save the output file bellow.
    
    $xPath = new DOMXPath($dom);

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    echo "criando planihla \n";

    $sheet->setCellValue('A1', 'ID');
    $sheet->setCellValue('B1', 'Title');
    $sheet->setCellValue('C1', 'Type');
    $sheet->setCellValue('D1', 'Author 1');
    $sheet->setCellValue('E1', 'Author 1 Institution');
    $sheet->setCellValue('F1', 'Author 2');
    $sheet->setCellValue('G1', 'Author 2 Institution');
    $sheet->setCellValue('H1', 'Author 3');
    $sheet->setCellValue('I1', 'Author 3 Institution');
    $sheet->setCellValue('J1', 'Author 4');
    $sheet->setCellValue('K1', 'Author 4 Institution');
    $sheet->setCellValue('L1', 'Author 5');
    $sheet->setCellValue('M1', 'Author 5 Institution');
    $sheet->setCellValue('N1', 'Author 6');
    $sheet->setCellValue('O1', 'Author 6 Institution');
    $sheet->setCellValue('P1', 'Author 7');
    $sheet->setCellValue('Q1', 'Author 7 Institution');
    $sheet->setCellValue('R1', 'Author 8');
    $sheet->setCellValue('S1', 'Author 8 Institution');
    $sheet->setCellValue('T1', 'Author 9');
    $sheet->setCellValue('U1', 'Author 9 Institution');

    $style = $sheet->getStyle('A1:V1');
    $font = $style->getFont();
    $font->setBold(true);

    echo "procurando cards \n";
    $cards = $xPath->query('.//a[contains(concat(" ",normalize-space(@class)," ")," paper-card ")]');
    //echo $cards;

     $row = 2;

     /** @var DOMNode $card */
     foreach ($cards as $card) {
       $id = $xPath->query('.//div//div[contains(concat(" ",normalize-space(@class)," ")," mr-sm ")]//div[contains(concat(" ",normalize-space(@class)," ")," volume-info ")]', $card);
       $titulo = $xPath->query('.//h4', $card);
       $type = $xPath->query('.//div//div[contains(concat(" ",normalize-space(@class)," ")," tags ")][contains(concat(" ",normalize-space(@class)," ")," mr-sm ")][not(self::node()[contains(concat(" ",normalize-space(@class)," ")," flex-row ")])]', $card);
       $authors = $xPath->query('.//div[contains(concat(" ",normalize-space(@class)," ")," authors ")]//span', $card);

      if ($id->length > 0) {
        // Acessa o primeiro elemento h2 e exibe seu texto
        $txt = trim($id[0]->textContent . PHP_EOL);
        $sheet->setCellValue('A' . $row, $txt);
        
      }else{
        echo "tá vindo tido vazio\n";
      }

      if ($titulo->length > 0) {
        // Acessa o primeiro elemento h2 e exibe seu texto
        $txt = trim($titulo[0]->textContent . PHP_EOL);
        $sheet->setCellValue('B' . $row, $txt);
        
      }else{
        echo "tá vindo tido vazio\n";
      }

      if ($type->length > 0) {
        // Acessa o primeiro elemento h2 e exibe seu texto
        $txt = $type[0]->textContent . PHP_EOL;
        $sheet->setCellValue('C' . $row, $txt);
        
      }else{
        echo "tá vindo tido vazio\n";
      }
      
       $contador = 0;
       foreach($authors as $author){
        if($contador < 9){
          $contador+=1;
        }
           
           $nometxt = $author->textContent . PHP_EOL;
           
           $inst = trim($author->getAttribute('title') . PHP_EOL);
           switch($contador){
               case 1:
                   $sheet->setCellValue('D' . $row, $nometxt);
                   $sheet->setCellValue('E'. $row, $inst);
                   break;
               case 2:
                   $sheet->setCellValue('F' . $row, $nometxt);
                   $sheet->setCellValue('G'. $row, $inst);
                   break;
               case 3:
                   $sheet->setCellValue('H' . $row, $nometxt);
                   $sheet->setCellValue('I'. $row, $inst);
                   break;
               case 4:
                   $sheet->setCellValue('J' . $row, $nometxt);
                   $sheet->setCellValue('K'. $row, $inst);
                   break;
               case 5:
                   $sheet->setCellValue('L' . $row, $nometxt);
                   $sheet->setCellValue('M'. $row, $inst);
                   echo "salvo\n";
                   break;
               case 6:
                   $sheet->setCellValue('N' . $row, $nometxt);
                   $sheet->setCellValue('O'. $row, $inst);
                   break;
               case 7:
                   $sheet->setCellValue('P' . $row, $nometxt);
                   $sheet->setCellValue('Q'. $row, $inst);
                   break;
               case 8:
                   $sheet->setCellValue('R' . $row, $nometxt);
                   $sheet->setCellValue('S'. $row, $inst);
                   break;
               case 9:
                   $sheet->setCellValue('T' . $row, $nometxt);
                   $sheet->setCellValue('U'. $row, $inst);
                   break;
               default:
                   echo "deu algo errado ae".$contador."\n";
           }
         }
       $row++;
      
     }
     $writer = new Xlsx($spreadsheet);
     $writer->save('assets/model.xlsx'); // Salvar o arquivo como informacoes.xlsx

     echo "Planilha criada com sucesso! \n";
   }

  


 
  

  
}
Main::run();