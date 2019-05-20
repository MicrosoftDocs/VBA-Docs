---
title: XmlMap.ExportXml method (Excel)
keywords: vbaxl10.chm754090
f1_keywords:
- vbaxl10.chm754090
ms.prod: excel
api_name:
- Excel.XmlMap.ExportXml
ms.assetid: ffb4e656-157e-e5f3-1ddd-314172ba5839
ms.date: 05/21/2019
localization_priority: Normal
---


# XmlMap.ExportXml method (Excel)

Exports the contents of cells mapped to the specified **XmlMap** object to a **String** variable.


## Syntax

_expression_.**ExportXml** (_Data_)

_expression_ A variable that represents an **[XmlMap](Excel.XmlMap.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Data_|Required| **String**|The variable to export the data to.|

## Return value

**[XlXmlExportResult](Excel.XlXmlExportResult.md)**


## Remarks

To export the contents of the mapped cells to an XML data file, use the **[Export](Excel.XmlMap.Export.md)** method.


## Example

The following example exports the contents of the cells mapped to the Contacts schema map to a variable named  `strContactData`.

```vb
Sub ExportToString() 
 Dim strContactData As String 
 
 ActiveWorkbook.XmlMaps("Contacts").ExportXml Data:=strContactData 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]