<?php

/**
 * @package civiexportexcel
 * @copyright Mathieu Lutfy (c) 2014
 *
 * Has a lot of duplication from CRM_CiviExportExcel_Utils_Report.
 * TODO: refactor later, if possible.
 */
class CRM_CiviExportExcel_Utils_SearchExport {

  /**
   * Generates a XLS 2007 file and forces the browser to download it.
   *
   * @param Array &$headers
   * @param Array &$columnTypes
   * @param Array &$rows
   */
  static function export2excel2007(&$headers, &$columnTypes, &$rows) {
    //Force a download and name the file using the current timestamp.
    $datetime = date('Ymd-Gi', $_SERVER['REQUEST_TIME']);

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="Export_' . $datetime . '.xlsx"');
    header("Content-Description: " . ts('CiviCRM export'));
    header("Content-Transfer-Encoding: binary");
    header("Cache-Control: no-cache, must-revalidate, post-check=0, pre-check=0");

    // always modified
    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');

    echo self::makeExcel($headers, $columnTypes, $rows);
    CRM_Utils_System::civiExit();
  }

  /**
   *
   * See @CRM_Report_Utils_SearchExport::export2excel2007().
   */
  static function makeExcel(&$headers, &$columnTypes, &$rows) {
    $config = CRM_Core_Config::singleton();
    $csv = '';

    // Generate an array with { 0=>A, 1=>B, 2=>C, ... }
    $foo = array(0 => '', 1 => 'A', 2 => 'B', 3 => 'C', 4 => 'D', 5 => 'E');
    $a = ord('A');
    $cells = array();

    for ($i = 0; $i < count($foo); $i++) {
      for ($j = 0; $j < 26; $j++) {
        $cells[$j + ($i * 26)] = $foo[$i] . chr($j + $a);
      }
    }

    include('PHPExcel/Classes/PHPExcel.php');
    $objPHPExcel = new PHPExcel();

    // Does magic things for date cells
    // https://phpexcel.codeplex.com/discussions/331005
    PHPExcel_Cell::setValueBinder(new PHPExcel_Cell_AdvancedValueBinder());

    // Set document properties
    $objPHPExcel->getProperties()
      ->setCreator("CiviCRM")
      ->setLastModifiedBy("CiviCRM")
      ->setTitle(ts('Export'))
      ->setSubject(ts('Export'))
      ->setDescription(ts('Export'));

    $sheet = $objPHPExcel->setActiveSheetIndex(0);
    $objPHPExcel->getActiveSheet()->setTitle('Export');

    // Add the column headers.
    $col = 0;
    $cpt = 1;

    foreach ($headers as $h) {
      try {
      $objPHPExcel->getActiveSheet()
        ->setCellValue($cells[$col] . $cpt, $h);
      }
      catch (Exception $e) {
        die(print_r($e, 1));
      }

      $col++;
    }

    // Add rows.
    $cpt = 2;

    // Convert the sql headers to civi types
    $columnHeaders = CRM_CiviExportExcel_Utils_SearchExport::sqlTypesToCivi($columnTypes);

    foreach ($rows as $row) {
      $displayRows = array();
      $col = 0;

      foreach ($columnTypes as $k => $v) {
        $value = CRM_Utils_Array::value($k, $row);

        if (! isset($value)) {
          $col++;
          continue;
        }

        // Remove HTML, unencode entities
        $value = html_entity_decode(strip_tags($value));

        $objPHPExcel->getActiveSheet()
          ->setCellValue($cells[$col] . $cpt, $value);

        // Cell formats
        if (CRM_Utils_Array::value('type', $columnHeaders[$k]) & CRM_Utils_Type::T_DATE) {
          $objPHPExcel->getActiveSheet()
            ->getStyle($cells[$col] . $cpt)
            ->getNumberFormat()
            ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDD);

          // Set autosize on date columns. 
          // We only do it for dates because we know they have a fixed width, unlike strings.
          // For eco-friendlyness, this should only be done once, perhaps when processing the headers initially
          $objPHPExcel->getActiveSheet()->getColumnDimension($cells[$col])->setAutoSize(true);
        }
        elseif (CRM_Utils_Array::value('type', $columnHeaders[$k]) & CRM_Utils_Type::T_MONEY) {
          $objPHPExcel->getActiveSheet()->getStyle($cells[$col])
            ->getNumberFormat()
            ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_GENERAL);
        }

        $col++;
      }

      $cpt++;
    }

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save('php://output');

    return ''; // FIXME
  }

  /**
   * For each header from the export, return the correct civicrm data type, if possible.
   * By default, the $sqlTypes are the mysql col definition (ex: varchar, int, etc),
   * but most columns are not correctly encoded. Ex: birth_date is shows as varchar(512).
   */
  static function sqlTypesToCivi(&$sqlTypes) {
    $headers = array();

    foreach ($sqlTypes as $key => $val) {
      if (strpos($val, 'date') !== FALSE) {
        $headers[$key] = array('type' => CRM_Utils_Type::T_DATE);
      }
      else {
        $headers[$key] = array('type' => CRM_Utils_Type::T_STRING);
      }
    }

    return $headers;
  }
}
