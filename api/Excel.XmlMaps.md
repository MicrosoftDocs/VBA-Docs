---
title: XmlMaps object (Excel)
keywords: vbaxl10.chm755072
f1_keywords:
- vbaxl10.chm755072
ms.prod: excel
api_name:
- Excel.XmlMaps
ms.assetid: 0cb16ec8-1120-0da3-508b-c1c9b0aa1701
ms.date: 04/03/2019
localization_priority: Normal
---


# XmlMaps object (Excel)

Represents the collection of **[XmlMap](Excel.XmlMap.md)** objects that have been added to a workbook.


## Example

Use the **Add** method to add an XML map to a workbook.

```vb
Sub AddXmlMap() 
 Dim strSchemaLocation As String 
 
 strSchemaLocation = "https://example.microsoft.com/schemas/CustomerData.xsd" 
 ActiveWorkbook.XmlMaps.Add strSchemaLocation, "Root" 
End Sub
```

## Methods

- [Add](Excel.XmlMaps.Add.md)

## Properties

- [Application](Excel.XmlMaps.Application.md)
- [Count](Excel.XmlMaps.Count.md)
- [Creator](Excel.XmlMaps.Creator.md)
- [Item](Excel.XmlMaps.Item.md)
- [Parent](Excel.XmlMaps.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]