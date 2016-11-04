<?php

/**
 * @package civiexportexcel
 * @copyright Mathieu Lutfy (c) 2014-2015
 */
class CRM_CiviExportExcel_Utils_Report extends CRM_Core_Page {

  /**
   * Generates a XLS 2007 file and forces the browser to download it.
   *
   * @param Object $form
   * @param Array &$rows
   *
   * See @CRM_Report_Utils_Report::export2csv().
   */
  static function export2excel2007(&$form, &$rows, &$stats) {
    //Force a download and name the file using the current timestamp.
    $datetime = date('Ymd-Gi', $_SERVER['REQUEST_TIME']);

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="Report_' . $datetime . '.xlsx"');
    header("Content-Description: " . ts('CiviCRM report'));
    header("Content-Transfer-Encoding: binary");
    header("Cache-Control: no-cache, must-revalidate, post-check=0, pre-check=0");

    // always modified
    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');

    self::generateFile($form, $rows, $stats);
    CRM_Utils_System::civiExit();
  }

  /**
   * Utility function for export2csv and CRM_Report_Form::endPostProcess
   * - make XLS file content and return as string.
   *
   * @param Object &$form CRM_Report_Form object.
   * @param Array &$rows Resulting rows from the report.
   * @param String Full path to the filename to write in (for mailing reports).
   *
   * See @CRM_Report_Utils_Report::makeCsv().
   */
  static function generateFile(&$form, &$rows, &$stats, $filename = 'php://output') {
    $config = CRM_Core_Config::singleton();
    $csv = '';

    // Generate an array with { 0=>A, 1=>B, 2=>C, ... }
    $foo = array(0 => '', 1 => 'A', 2 => 'B', 3 => 'C', 4 => 'D', 5 => 'E', 6 => 'F', 7 => 'G', 8 => 'H', 9 => 'I', 10 => 'J', 11 => 'K', 12 => 'L', 13 => 'M');
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

    // FIXME Set the locale of the XLS file
    // might not really be necessary (concerns mostly functions? not dates?)
    // $validLocale = PHPExcel_Settings::setLocale('fr');

    // Set document properties
    $objPHPExcel->getProperties()
      ->setCreator("CiviCRM")
      ->setLastModifiedBy("CiviCRM")
      ->setTitle(ts('Report'))
      ->setSubject(ts('Report'))
      ->setDescription(ts('Report'));

    $sheet = $objPHPExcel->setActiveSheetIndex(0);
    $objPHPExcel->getActiveSheet()->setTitle('Report');

    // Add headers if this is the first row.
    $columnHeaders = array_keys($form->_columnHeaders);

    // Replace internal header names with friendly ones, where available.
    foreach ($columnHeaders as $header) {
      if (isset($form->_columnHeaders[$header])) {
        $headers[] = html_entity_decode(strip_tags($form->_columnHeaders[$header]['title']));
      }
    }

    // Add the column headers.
    $col = 0;
    $cpt = 1;

    foreach ($headers as $h) {
      $objPHPExcel->getActiveSheet()
        ->setCellValue($cells[$col] . $cpt, $h);

      $col++;
    }

    // Add rows.
    $cpt = 2;

    foreach ($rows as $row) {
      $displayRows = array();
      $col = 0;

      foreach ($columnHeaders as $k => $v) {
        $value = CRM_Utils_Array::value($v, $row);

        if (! isset($value)) {
          $col++;
          continue;
        }

        // Remove HTML, unencode entities
        $value = html_entity_decode(strip_tags($value));

        // Data transformation before adding it to the cell
        if (CRM_Utils_Array::value('type', $form->_columnHeaders[$v]) & CRM_Utils_Type::T_DATE) {
          $group_by = CRM_Utils_Array::value('group_by', $form->_columnHeaders[$v]);

          if ($group_by == 'MONTH' || $group_by == 'QUARTER') {
            $value = CRM_Utils_Date::customFormat($value, $config->dateformatPartial);
          }
          elseif ($group_by == 'YEAR') {
            $value = CRM_Utils_Date::customFormat($value, $config->dateformatYear);
          }
          else {
            $value = CRM_Utils_Date::customFormat($value, '%Y-%m-%d');
          }
        }

        $objPHPExcel->getActiveSheet()
          ->setCellValue($cells[$col] . $cpt, $value);

        // Cell formats
        if (CRM_Utils_Array::value('type', $form->_columnHeaders[$v]) & CRM_Utils_Type::T_DATE) {
          $objPHPExcel->getActiveSheet()
            ->getStyle($cells[$col] . $cpt)
            ->getNumberFormat()
            ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDD);

          // Set autosize on date columns. 
          // We only do it for dates because we know they have a fixed width, unlike strings.
          // For eco-friendlyness, this should only be done once, perhaps when processing the headers initially
          $objPHPExcel->getActiveSheet()->getColumnDimension($cells[$col])->setAutoSize(true);
        }
        elseif (CRM_Utils_Array::value('type', $form->_columnHeaders[$v]) & CRM_Utils_Type::T_MONEY) {
          $objPHPExcel->getActiveSheet()->getStyle($cells[$col])
            ->getNumberFormat()
            ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_GENERAL);
        }

        $col++;
      }

      $cpt++;
    }

    // Add report statistics on a separate Excel sheet.
    if (! empty($stats) && ! empty($stats['counts'])) {
      $cpt = 1;

      $objWorkSheet = $objPHPExcel->createSheet(1);
      $objWorkSheet->setTitle(ts('Statistics'));

      foreach ($stats['counts'] as $key => $val) {
        $objWorkSheet
          ->setCellValue('A' . $cpt, $val['title'])
          ->setCellValue('B' . $cpt, $val['value']);

        $cpt++;
      }

      $objWorkSheet->getColumnDimension('A')->setWidth(30);
    }

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save($filename);

    return ''; // FIXME
  }
}
