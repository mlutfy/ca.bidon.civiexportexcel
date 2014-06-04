CiviCRM export to Excel
=======================

Experimental extension!

This extension adds the possibility to export directly into the MS Excel
format from CiviReports, instead of CSV (less fiddling, easier to use).

This extension uses the PHPExcel library. See the "License" section below
for more information.

For discussion, see: http://forum.civicrm.org/index.php/topic,32954.0.html
See also the "Todo" section for a general roadmap.

To download the latest version of this module:
https://github.com/mlutfy/ca.bidon.civiexportexcel

Requirements
============

- CiviCRM >= 4.4 (previous versions untested)
- patching CiviCRM core code.

Installation
============

1- Download this extension and unpack it in your 'extensions' directory.
   You may need to create it if it does not already exist, and configure
   the correct path in CiviCRM -> Administer -> System -> Directories.

2- Enable the extension from CiviCRM -> Administer -> System -> Extensions.

3- Apply a patch to CiviCRM core, otherwise the "Export to Excel" button will
   not be available from CiviReports. See file: civiexportexcel-core.patch
   included in the main directory of this extension.

Support
=======

Please post bug reports in the issue tracker of this project on github:
https://github.com/mlutfy/ca.bidon.civiexportexcel/issues

This is a community contributed extension written thanks to the financial
support of organisations using it, as well as the very helpful and collaborative
CiviCRM community.

If you appreciate this module, please consider donating 10$ to the CiviCRM project:
http://civicrm.org/participate/support-civicrm

While I do my best to provide volunteer support for this extension, please
consider financially contributing to support or development of this extension
if you can.
http://www.bidon.ca/en/paypal

Todo
====

* Support report e-mails (return the output of the xls, do not output directly as download).
* Propose a new hook to CiviCRM so that we do not need to patch core.
* Apply new hook to Search export as well (in the "select fields" step, add a "format" option that calls the hook?)
* Add column headers in xls export.
* Add OpenDocument (LibreOffice) support.
* Add admin settings form to enable excel/opendocument formats?

License
=======

(C) 2014 Mathieu Lutfy <mathieu@bidon.ca>

Distributed under the terms of the GNU Affero General public license (AGPL).
See LICENSE.txt for details.

This extension includes PHPExcel:

Version 1.8.0, 2014-03-02
Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt
