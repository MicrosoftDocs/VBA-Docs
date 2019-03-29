---
title: OLEFormat object (Excel)
keywords: vbaxl10.chm631072
f1_keywords:
- vbaxl10.chm631072
ms.prod: excel
api_name:
- Excel.OLEFormat
ms.assetid: 96ee06d8-e922-c48c-4406-bb2f5cbaa02a
ms.date: 03/30/2019
localization_priority: Normal
---


# OLEFormat object (Excel)

Contains OLE object properties.


## Remarks

If the **[Shape](Excel.Shape.md)** object doesn't represent a linked or embedded object, the **[OLEFormat](Excel.Shape.OLEFormat.md)** property of the **Shape** object fails.


## Example

Use the **OLEFormat** property to return the **OLEFormat** object. The following example activates an OLE object in the **[Shapes](Excel.Shapes.md)** collection.

```vb
Worksheets(1).Shapes(1).OLEFormat.Activate
```

## Methods


- [Activate](Excel.OLEFormat.Activate.md)
- [Verb](Excel.OLEFormat.Verb.md)

## Properties

- [Application](Excel.OLEFormat.Application.md)
- [Creator](Excel.OLEFormat.Creator.md)
- [Object](Excel.OLEFormat.Object.md)
- [Parent](Excel.OLEFormat.Parent.md)
- [progID](Excel.OLEFormat.progID.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]