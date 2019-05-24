---
title: Workbook.AfterXmlImport event (Excel)
keywords: vbaxl10.chm503098
f1_keywords:
- vbaxl10.chm503098
ms.prod: excel
api_name:
- Excel.Workbook.AfterXmlImport
ms.assetid: b43adf53-6b67-6127-e69d-6ea05f68b7f6
ms.date: 05/25/2019
localization_priority: Normal
---


# Workbook.AfterXmlImport event (Excel)

Occurs after an existing XML data connection is refreshed or after new XML data is imported into the specified Microsoft Excel workbook.


## Syntax

_expression_.**AfterXmlImport** (_Map_, _IsRefresh_, _Result_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Map_|Required| **[XmlMap](Excel.XmlMap.md)**|The XML map that will be used to import data.|
| _IsRefresh_|Required| **Boolean**| **True** if the event was triggered by refreshing an existing connection to XML data; **False** if the event was triggered by importing from a different data source.|
| _Result_|Required| **[XlXmlImportResult](Excel.XlXmlImportResult.md)**|Indicates the results of the refresh or import operation.|

## Return value

**Nothing**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]