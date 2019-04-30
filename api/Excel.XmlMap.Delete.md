---
title: XmlMap.Delete method (Excel)
keywords: vbaxl10.chm754086
f1_keywords:
- vbaxl10.chm754086
ms.prod: excel
api_name:
- Excel.XmlMap.Delete
ms.assetid: 8acde534-c465-029a-635a-38f63c5f4013
ms.date: 06/08/2017
localization_priority: Normal
---


# XmlMap.Delete method (Excel)

Removes the specified XML map from the workbook.


## Syntax

_expression_.**Delete**

_expression_ A variable that represents an **[XmlMap](Excel.XmlMap.md)** object.


## Remarks

Deleting the XML map will convert all the XML Lists to generic Lists and remove all the single-cell mappings (with the data still remaining). In addition, the **XmlMap** object will be removed from the **[XmlMaps](Excel.XmlMaps.md)** collection. The map and schema information will be removed from the workbook (it will no longer be persisted in the XLS file and XMLSS). Any references to the deleted object become invalid.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]