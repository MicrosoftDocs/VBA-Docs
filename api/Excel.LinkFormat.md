---
title: LinkFormat object (Excel)
keywords: vbaxl10.chm633072
f1_keywords:
- vbaxl10.chm633072
ms.prod: excel
api_name:
- Excel.LinkFormat
ms.assetid: 3d8085bf-c113-7cbe-871b-01f3b6017824
ms.date: 03/30/2019
localization_priority: Normal
---


# LinkFormat object (Excel)

Contains linked OLE object properties.


## Remarks

If the **[Shape](Excel.Shape.md)** object doesn't represent a linked object, the **[LinkFormat](Excel.Shape.LinkFormat.md)** property of the **Shape** object fails.


## Example

Use the **LinkFormat** property to return the **LinkFormat** object. 

The following example updates an OLE object in the **[Shapes](Excel.Shapes.md)** collection.

```vb
Worksheets(1).Shapes(1).LinkFormat.Update
```

## Methods

- [Update](Excel.LinkFormat.Update.md)

## Properties

- [Application](Excel.LinkFormat.Application.md)
- [AutoUpdate](Excel.LinkFormat.AutoUpdate.md)
- [Creator](Excel.LinkFormat.Creator.md)
- [Locked](Excel.LinkFormat.Locked.md)
- [Parent](Excel.LinkFormat.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]