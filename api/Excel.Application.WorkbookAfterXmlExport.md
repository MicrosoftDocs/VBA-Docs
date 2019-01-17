---
title: Application.WorkbookAfterXmlExport Event (Excel)
keywords: vbaxl10.chm504101
f1_keywords:
- vbaxl10.chm504101
ms.prod: excel
api_name:
- Excel.Application.WorkbookAfterXmlExport
ms.assetid: 9d542c67-4244-d018-4db6-3584f0caec7c
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WorkbookAfterXmlExport Event (Excel)

Occurs after Microsoft Excel saves or exports XML data from the specified workbook.


## Syntax

_expression_. `WorkbookAfterXmlExport`( `_Wb_` , `_Map_` , `_Url_` , `_Result_` )

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The target workbook.|
| _Map_|Required| **[XmlMap](Excel.XmlMap.md)**|The XML map that was used to save or export data.|
| _Url_|Required| **String**|The location of the XML file that was exported.|
| _Result_|Required| **[xlXmlExportResult](Excel.XlXmlExportResult.md)**| Indicates the results of the save or export operation.|

## Return value

Nothing


## Remarks



| **xlXmlExportResult** can be one of the following **xlXmlExportResult** constants|
| **xlXmlExportSuccess**. The XML data file was successfully exported.|
| **xlXmlExportValidationFailed**. The contents of the XML data file do not match the specified schema map.|

Use the  **[AfterXmlExport](Excel.Workbook.AfterXmlExport.md)** event if you want to perform an operation after XML data has been exported from a particular workbook.


## See also


[Workbook Object](Excel.Workbook.md)
[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]