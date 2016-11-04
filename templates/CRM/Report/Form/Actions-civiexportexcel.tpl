{* We need this reference for the js injection below *}
{assign var=csv value="_qf_"|cat:$form.formName|cat:"_submit_csv"}

{* The nbps; are a mimic of what other buttons do in templates/CRM/Report/Form/Actions.tpl *}
{assign var=excel value="_qf_"|cat:$form.formName|cat:"_submit_excel"}
{$form.$excel.html}&nbsp;&nbsp;

{literal}
  <script>
    CRM.$(function($) {
      var form_id = '{/literal}{$form.$excel.id}{literal}';

      {/literal}
        {* CiviCRM 4.6 *}
        {if $form.$csv.id}
          {literal}
            var $dest = $('input#{/literal}{$form.$csv.id}{literal}').parent();
            $('input#' + form_id).appendTo($dest);
          {/literal}
        {else}
          {* CiviCRM 4.7+ *}
          {literal}
            if ($('.crm-report-field-form-block .crm-submit-buttons').size() > 0) {
              $('input#' + form_id).appendTo('.crm-report-field-form-block .crm-submit-buttons');
            }
            else {
              // Do not show the button when running in a dashlet
              // FIXME: we should probably just not add the HTML in the first place.
              $('input#' + form_id).hide();
            }
          {/literal}
        {/if}
      {literal}
    });
  </script>
{/literal}
