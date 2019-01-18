---
title: Application.WorkbookAfterXmlImport Event (Excel)
keywords: vbaxl10.chm504099
f1_keywords:
- vbaxl10.chm504099
ms.prod: excel
api_name:
- Excel.Application.WorkbookAfterXmlImport
ms.assetid: a58cc327-3776-fe5b-68d4-406269f30379
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WorkbookAfterXmlImport Event (Excel)

Occurs after an existing XML data connection is refreshed, or new XML data is imported into any open Microsoft Excel workbook.


## Syntax

_expression_. `WorkbookAfterXmlImport`( `_Wb_` , `_Map_` , `_IsRefresh_` , `_Result_` )

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The target workbook.|
| _Map_|Required| **[XmlMap](Excel.XmlMap.md)**|The XML map that was used to import data.|
| _IsRefresh_|Required| **Boolean**| **True** if the event was triggered by refreshing an existing connection to XML data, **False** if a new mapping was created.|
| _Result_|Required| **[xlXmlImportResult](Excel.XlXmlImportResult.md)**|Indicates the results of the refresh or import operation.|

## Return value

Nothing


## Remarks



| **xlXmlImportResult** can be one of the following **xlXmlImportResult** constants|
| **xlXmlImportElementsTruncated**. The contents of the specified XML data file have been truncated because the XML data file is too large for the worksheet.|
| **xlXmlImportSuccess**. The XML data file was successfully imported.|
| **xlXmlImportValidationFailed**. The contents of the XML data file do not match the specified schema map.|

Use the  **[AfterXmlImport](Excel.Workbook.AfterXmlImport.md)** event if you want to perform an operation after XML data has been imported into a particular workbook.


## See also


[Workbook Object](Excel.Workbook.md)
[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]