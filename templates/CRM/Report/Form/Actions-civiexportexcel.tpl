{* We need this reference for the js injection below *}
{assign var=csv value="_qf_"|cat:$form.formName|cat:"_submit_csv"}

{* The nbps; are a mimic of what other buttons do in templates/CRM/Report/Form/Actions.tpl *}
{assign var=excel value="_qf_"|cat:$form.formName|cat:"_submit_excel"}
{$form.$excel.html}&nbsp;&nbsp;

{literal}
  <script>
    cj(function() {
      var form_id = '{/literal}{$form.$excel.id}{literal}';
      var $dest = cj('input#{/literal}{$form.$csv.id}{literal}').parent();
      cj('input#' + form_id).appendTo($dest);
    });
  </script>
{/literal}
