<?php

require_once 'civiexportexcel.civix.php';

/**
 * Implementation of hook_civicrm_config
 */
function civiexportexcel_civicrm_config(&$config) {
  _civiexportexcel_civix_civicrm_config($config);
}

/**
 * Implementation of hook_civicrm_xmlMenu
 *
 * @param $files array(string)
 */
function civiexportexcel_civicrm_xmlMenu(&$files) {
  _civiexportexcel_civix_civicrm_xmlMenu($files);
}

/**
 * Implementation of hook_civicrm_install
 */
function civiexportexcel_civicrm_install() {
  return _civiexportexcel_civix_civicrm_install();
}

/**
 * Implementation of hook_civicrm_uninstall
 */
function civiexportexcel_civicrm_uninstall() {
  return _civiexportexcel_civix_civicrm_uninstall();
}

/**
 * Implementation of hook_civicrm_enable
 */
function civiexportexcel_civicrm_enable() {
  return _civiexportexcel_civix_civicrm_enable();
}

/**
 * Implementation of hook_civicrm_disable
 */
function civiexportexcel_civicrm_disable() {
  return _civiexportexcel_civix_civicrm_disable();
}

/**
 * Implementation of hook_civicrm_upgrade
 *
 * @param $op string, the type of operation being performed; 'check' or 'enqueue'
 * @param $queue CRM_Queue_Queue, (for 'enqueue') the modifiable list of pending up upgrade tasks
 *
 * @return mixed  based on op. for 'check', returns array(boolean) (TRUE if upgrades are pending)
 *                for 'enqueue', returns void
 */
function civiexportexcel_civicrm_upgrade($op, CRM_Queue_Queue $queue = NULL) {
  return _civiexportexcel_civix_civicrm_upgrade($op, $queue);
}

/**
 * Implementation of hook_civicrm_managed
 *
 * Generate a list of entities to create/deactivate/delete when this module
 * is installed, disabled, uninstalled.
 */
function civiexportexcel_civicrm_managed(&$entities) {
  return _civiexportexcel_civix_civicrm_managed($entities);
}

/**
 * Implementation of hook_civicrm_buildForm().
 *
 * Used to add a 'Export to Excel' button in the Report forms.
 */
function civiexportexcel_civicrm_buildForm($formName, &$form) {
  // Reports extend the CRM_Report_Form class.
  // We use that to check whether we should inject the Excel export buttons.
  if (is_subclass_of($form, 'CRM_Report_Form')) {
    $smarty = CRM_Core_Smarty::singleton();
    $vars = $smarty->get_template_vars();

    $form->_excelButtonName = $form->getButtonName('submit', 'excel');

    $label = (! empty($vars['instanceId']) ? ts('Export to Excel') : ts('Preview Excel'));
    $form->addElement('submit', $form->_excelButtonName, $label);

    CRM_Core_Region::instance('page-body')->add(array(
      'template' => 'CRM/Report/Form/Actions-civiexportexcel.tpl',
    ));

    // This hook also gets called when we click on a submit button,
    // so we can handle that part here too.
    $buttonName = $form->controller->getButtonName();

    $output = CRM_Utils_Request::retrieve('output', 'String', CRM_Core_DAO::$_nullObject);

    if ($form->_excelButtonName == $buttonName || $output == 'excel2007') {
      $form->assign('printOnly', TRUE);
      $printOnly = TRUE;
      $form->assign('outputMode', 'excel2007');

      // FIXME: this duplicates part of CRM_Report_Form::postProcess()
      // since we do not have a place to hook into, we hi-jack the form process
      // before it gets into postProcess.

      // get ready with post process params
      $form->beginPostProcess();

      // build query
      $sql = $form->buildQuery(FALSE);

      // build array of result based on column headers. This method also allows
      // modifying column headers before using it to build result set i.e $rows.
      $rows = array();
      $form->buildRows($sql, $rows);

      // format result set.
      // This seems to cause more problems than it fixes.
      // $form->formatDisplay($rows);

      // Show stats on a second Excel page.
      $stats = $form->statistics($rows);

      // assign variables to templates
      $form->doTemplateAssignment($rows);
      // FIXME: END.

      CRM_CiviExportExcel_Utils_Report::export2excel2007($form, $rows, $stats);
    }
  }
}

/**
 * Implements hook_civicrm_export().
 *
 * Called mostly to export search results.
 */
function civiexportexcel_civicrm_export($exportTempTable, $headerRows, $sqlColumns, $exportMode) {
  $writeHeader = true;

  $rows = array();

  $query = "SELECT * FROM $exportTempTable";
  $dao = CRM_Core_DAO::executeQuery($query);

  while ($dao->fetch()) {
    $row = array();
    foreach ($sqlColumns as $column => $dontCare) {
      $row[$column] = $dao->$column;
    }

    $rows[] = $row;
  }

  $dao->free();

  CRM_CiviExportExcel_Utils_SearchExport::export2excel2007($headerRows, $sqlColumns, $rows);
}

/**
 * Implements hook_civicrm_alterMailParams().
 *
 * Intercepts outgoing report emails, in order to attach the
 * excel2007 version of the report.
 *
 * TODO: we should really propose a patch to CRM_Report_Form::endPostProcess().
 */
function civiexportexcel_attach_to_email(&$form, &$rows, &$attachments) {
  $config = CRM_Core_Config::singleton();

  $filename = 'CiviReport.xlsx';
  $fullname = $config->templateCompileDir . CRM_Utils_File::makeFileName($filename);

  CRM_CiviExportExcel_Utils_Report::generateFile($form, $rows, $fullname);

  $attachments[] = array(
    'fullPath' => $fullname,
    'mime_type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'cleanName' => $filename,
  );
}
