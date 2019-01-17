---
title: Application.WorkbookBeforeXmlExport Event (Excel)
keywords: vbaxl10.chm504100
f1_keywords:
- vbaxl10.chm504100
ms.prod: excel
api_name:
- Excel.Application.WorkbookBeforeXmlExport
ms.assetid: 2c228d28-2d42-40b0-ee36-214bc720d78a
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WorkbookBeforeXmlExport Event (Excel)

Occurs before Microsoft Excel saves or exports XML data from the specified workbook.


## Syntax

_expression_. `WorkbookBeforeXmlExport`( `_Wb_` , `_Map_` , `_Url_` , `_Cancel_` )

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The target workbook.|
| _Map_|Required| **[XmlMap](Excel.XmlMap.md)**|The XML map that will be used to save or export data.|
| _Url_|Required| **String**|The location of the XML file to be exported.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the save or export operation.|

## Return value

Nothing


## Remarks

Use the  **[BeforeXmlExport](Excel.Workbook.BeforeXmlImport.md)** event if you want to capture XML data that is being exported or saved from a particular workbook.


## See also


[Workbook Object](Excel.Workbook.md)
[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]