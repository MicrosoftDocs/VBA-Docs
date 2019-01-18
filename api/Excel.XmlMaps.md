---
title: XmlMaps object (Excel)
keywords: vbaxl10.chm755072
f1_keywords:
- vbaxl10.chm755072
ms.prod: excel
api_name:
- Excel.XmlMaps
ms.assetid: 0cb16ec8-1120-0da3-508b-c1c9b0aa1701
ms.date: 06/08/2017
localization_priority: Normal
---


# XmlMaps object (Excel)

Represents the collection of  **[XmlMap](Excel.XmlMap.md)** objects that have been added to a workbook.


## Example

Use the  **[Add](Excel.XmlMaps.Add.md)** method to add an XML map to a workbook.


```vb
Sub AddXmlMap() 
 Dim strSchemaLocation As String 
 
 strSchemaLocation = "https://example.microsoft.com/schemas/CustomerData.xsd" 
 ActiveWorkbook.XmlMaps.Add strSchemaLocation, "Root" 
End Sub
```


## See also


[Excel Object Model Reference](./overview/Excel/object-model.md)


