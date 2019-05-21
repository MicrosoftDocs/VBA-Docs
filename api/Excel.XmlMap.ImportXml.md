---
title: XmlMap.ImportXml method (Excel)
keywords: vbaxl10.chm754088
f1_keywords:
- vbaxl10.chm754088
ms.prod: excel
api_name:
- Excel.XmlMap.ImportXml
ms.assetid: 07db07d3-cd0f-08fe-3463-04ca72d084d1
ms.date: 05/21/2019
localization_priority: Normal
---


# XmlMap.ImportXml method (Excel)

Imports XML data from a **String** variable into cells that have been mapped to the specified **XmlMap** object.


## Syntax

_expression_.**ImportXml** (_XmlData_, _Overwrite_)

_expression_ A variable that represents an **[XmlMap](Excel.XmlMap.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _XmlData_|Required| **String**|The string that contains the XML data to import.|
| _Overwrite_|Optional| **Variant**|Specifies whether to overwrite the contents of cells that are currently mapped to the specified XML map. Set to **True** to overwrite the cells; set to **False** to append the data to the existing range.<br/><br/>If this parameter is not specified, the current value of the **[AppendOnImport](Excel.XmlMap.AppendOnImport.md)** property of the XML map determines whether the contents of cells are overwritten.|

## Return value

**[XlXmlImportResult](Excel.XlXmlImportResult.md)**


## Remarks

To import the contents of an XML data file into cells mapped to a specific schema map, use the **[Import](Excel.XmlMap.Import.md)** method of the **XmlMap** object.

If either of the following conditions is **True**, a run-time error occurs. If more than one condition is **True**, Excel returns a run-time error for the most severe (they are listed with the most severe listed first):

- If the XML data contains syntactical errors.   
- If the import is cancelled because not all the data could fit on the worksheet.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]