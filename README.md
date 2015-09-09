CiviCRM export to Excel
=======================

This extension adds the possibility to export directly into the MS Excel
format from CiviReports and Search results, instead of CSV (less fiddling,
easier to use).

This extension uses the PHPExcel library. See the "License" section below
for more information (LGPL v2).

For discussion, see: http://forum.civicrm.org/index.php/topic,32954.0.html  
See also the "Todo" section for a general roadmap.

To download the latest version of this module:  
https://github.com/mlutfy/ca.bidon.civiexportexcel

This extension was sponsored by:  
Projet Montréal <http://projetmontreal.org>  
Development and Peace <https://www.devp.org>  
Coop SymbioTIC <https://www.symbiotic.coop>

Warnings
========

* This extension does not run the buildACLClause() function, meaning that you may have deleted contacts show up in some reports. If you are using ACLs in general, this can also cause important issues.

Requirements
============

- CiviCRM >= 4.4 (previous versions untested)

Installation
============

Install as any other regular CiviCRM extension:

1- Download this extension and unpack it in your 'extensions' directory.
   You may need to create it if it does not already exist, and configure
   the correct path in CiviCRM -> Administer -> System -> Directories.

2- Enable the extension from CiviCRM -> Administer -> System -> Extensions.

3- If you wish to send emails with the report as an Excel attachment,
   you must apply the patch in civiexportexcel-core-mail.patch.

Report mails
============

To send report e-mails in Excel2007 format, use: "format=excel2007" in
the "Scheduled Jobs" settings.

Support
=======

Please post bug reports in the issue tracker of this project on github:  
https://github.com/mlutfy/ca.bidon.civiexportexcel/issues

For general questions and support, please post on the "Extensions" forum:  
http://forum.civicrm.org/index.php/board,57.0.html

This is a community contributed extension written thanks to the financial
support of organisations using it, as well as the very helpful and collaborative
CiviCRM community.

If you appreciate this module, please consider donating 10$ to the CiviCRM project:  
http://civicrm.org/participate/support-civicrm

While I do my best to provide volunteer support for this extension, please
consider financially contributing to support or development of this extension
if you can.

Commercial support via Coop SymbioTIC:  
<https://www.symbiotic.coop>

Or you can send me the equivalent of a beer:  
<https://www.bidon.ca/en/paypal>

Todo
====

* Propose a new hook to CiviCRM for a cleaner postProcess implementation (incl. mail).
* Add OpenDocument (LibreOffice) support.
* Add admin settings form to enable excel/opendocument formats?
* Compare performance with tinybutstrong/tbs library?

License
=======

(C) 2014-2015 Mathieu Lutfy <mathieu@bidon.ca>

Distributed under the terms of the GNU Affero General public license (AGPL).
See LICENSE.txt for details.

This extension includes PHPExcel:

Version 1.8.0, 2014-03-02
Copyright (c) 2006 - 2013 PHPExcel (http://www.codeplex.com/PHPExcel)
http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt
