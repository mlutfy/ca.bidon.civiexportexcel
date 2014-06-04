<?php

/**
 * @package civireportexcel
 * @copyright Mathieu Lutfy (c) 2014
 */
class CRM_CiviExportExcel_Utils_Report {

  /**
   * Generates a XLS 2007 file and forces the browser to download it.
   *
   * @param Object $form
   * @param Array &$rows
   *
   * See @CRM_Report_Utils_Report::export2csv().
   */
  static function export2excel2007(&$form, &$rows) {
    //Force a download and name the file using the current timestamp.
    $datetime = date('Ymd-Gi', $_SERVER['REQUEST_TIME']);

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="Report_' . $datetime . '.xlsx"');
    header("Content-Description: " . ts('CiviCRM report'));
    header("Content-Transfer-Encoding: binary");
    header("Cache-Control: no-cache, must-revalidate, post-check=0, pre-check=0");

    // always modified
    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');

    echo self::makeCsv($form, $rows);
    CRM_Utils_System::civiExit();
  }

  /**
   * Utility function for export2csv and CRM_Report_Form::endPostProcess
   * - make XLS file content and return as string.
   *
   * FIXME: return as string, not output directly.
   *
   * See @CRM_Report_Utils_Report::makeCsv().
   */
  static function makeCsv(&$form, &$rows) {
    $config = CRM_Core_Config::singleton();
    $csv = '';

    // Generate an array with { 0=>A, 1=>B, 2=>C, ... }
    $a = ord('A');
    $cells = array();

    for ($i = 0; $i < 26; $i++) {
      $cells[$i] = chr($i + $a);
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

    $cpt = 1;
    $cell = 0; // starts at A1 using $cells

    // Add headers if this is the first row.
    $columnHeaders = array_keys($form->_columnHeaders);

    // Replace internal header names with friendly ones, where available.
    foreach ($columnHeaders as $header) {
      if (isset($form->_columnHeaders[$header])) {
        $headers[] = '"' . html_entity_decode(strip_tags($form->_columnHeaders[$header]['title'])) . '"';
      }
    }

    // Add the headers. FIXME
    // $csv .= implode($config->fieldSeparator, $headers) . "\r\n";

    foreach ($rows as $row) {
      $displayRows = array();
      $col = 0;

      foreach ($columnHeaders as $k => $v) {
        $value = CRM_Utils_Array::value($v, $row);

        if (! isset($value)) {
          $col++;
          continue;
        }

        if (CRM_Utils_Array::value('type', $form->_columnHeaders[$v]) & 4) {
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
        elseif (CRM_Utils_Array::value('type', $form->_columnHeaders[$v]) == 1024) {
          $value = CRM_Utils_Money::format($value);
        }

        $objPHPExcel->getActiveSheet()
          ->setCellValue($cells[$col] . $cpt, $value);

        if (CRM_Utils_Array::value('type', $form->_columnHeaders[$v]) & 4) {
          $objPHPExcel->getActiveSheet()
            ->getStyle($cells[$col] . $cpt)
            ->getNumberFormat()
            ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDDSLASH);

          // Set autosize on date columns. 
          // We only do it for dates because we know they have a fixed width, unlike strings.
          // For eco-friendlyness, this should only be done once, perhaps when processing the headers initially
          $objPHPExcel->getActiveSheet()->getColumnDimension($cells[$col])->setAutoSize(true);
        }

        $col++;
      }

      $cpt++;
    }

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save('php://output');

    return ''; // FIXME
  }
}
