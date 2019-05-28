---
title: Workbook.BeforeXmlExport event (Excel)
keywords: vbaxl10.chm503099
f1_keywords:
- vbaxl10.chm503099
ms.prod: excel
api_name:
- Excel.Workbook.BeforeXmlExport
ms.assetid: ee2af5de-e52f-9434-aa7c-5dc9bb102d1b
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.BeforeXmlExport event (Excel)

Occurs before Microsoft Excel saves or exports XML data from the specified workbook.


## Syntax

_expression_.**BeforeXmlExport** (_Map_, _Url_, _Cancel_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Map_|Required| **[XmlMap](Excel.XmlMap.md)**|The XML map that is used to save or export data.|
| _Url_|Required| **String**|The location where you want to export the resulting XML file.|
| _Cancel_|Required| **Boolean**|Set to **True** to cancel the save or export operation.|

## Return value

**Nothing**


## Remarks

This event does not occur when you are saving to the XML Spreadsheet file format.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]