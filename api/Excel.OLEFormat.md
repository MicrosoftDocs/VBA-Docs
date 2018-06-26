---
title: OLEFormat Object (Excel)
keywords: vbaxl10.chm631072
f1_keywords:
- vbaxl10.chm631072
ms.prod: excel
api_name:
- Excel.OLEFormat
ms.assetid: 96ee06d8-e922-c48c-4406-bb2f5cbaa02a
ms.date: 06/08/2017
---


# OLEFormat Object (Excel)

Contains OLE object properties.


## Remarks

If the  **[Shape](Excel.Shape.md)** object doesn't represent a linked or embedded object, the **[OLEFormat](Excel.Shape.OLEFormat.md)** property fails.


## Example

Use the  **OLEFormat** property to return the **OLEFormat** object. The following example activates an OLE object in the **[Shapes](Excel.Shapes.md)** collection.


```vb
Worksheets(1).Shapes(1).OLEFormat.Activate
```


## See also


[Excel Object Model Reference](./overview/object-model-excel-vba-reference.md)


