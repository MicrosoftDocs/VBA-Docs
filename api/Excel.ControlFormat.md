---
title: ControlFormat object (Excel)
keywords: vbaxl10.chm629072
f1_keywords:
- vbaxl10.chm629072
ms.prod: excel
api_name:
- Excel.ControlFormat
ms.assetid: fafc6e6b-641c-2179-0789-d86c2718b3c0
ms.date: 03/29/2019
localization_priority: Normal
---


# ControlFormat object (Excel)

Contains Microsoft Excel control properties.


## Example

Use the **[ControlFormat](Excel.Shape.ControlFormat.md)** property of the **Shape** object to return a **ControlFormat** object. The following example sets the fill range for a list box control on worksheet one.

> [!NOTE] 
> If the shape isn't a control, the **ControlFormat** property fails; and if the control isn't a list box, the **ListFillRange** property fails.

```vb
Worksheets(1).Shapes(1).ControlFormat.ListFillRange = "A1:A10"
```

## Methods

- [AddItem](Excel.ControlFormat.AddItem.md)
- [List](Excel.ControlFormat.List.md)
- [RemoveAllItems](Excel.ControlFormat.RemoveAllItems.md)
- [RemoveItem](Excel.ControlFormat.RemoveItem.md)

## Properties

- [Application](Excel.ControlFormat.Application.md)
- [Creator](Excel.ControlFormat.Creator.md)
- [DropDownLines](Excel.ControlFormat.DropDownLines.md)
- [Enabled](Excel.ControlFormat.Enabled.md)
- [LargeChange](Excel.ControlFormat.LargeChange.md)
- [LinkedCell](Excel.ControlFormat.LinkedCell.md)
- [ListCount](Excel.ControlFormat.ListCount.md)
- [ListFillRange](Excel.ControlFormat.ListFillRange.md)
- [ListIndex](Excel.ControlFormat.ListIndex.md)
- [LockedText](Excel.ControlFormat.LockedText.md)
- [Max](Excel.ControlFormat.Max.md)
- [Min](Excel.ControlFormat.Min.md)
- [MultiSelect](Excel.ControlFormat.MultiSelect.md)
- [Parent](Excel.ControlFormat.Parent.md)
- [PrintObject](Excel.ControlFormat.PrintObject.md)
- [SmallChange](Excel.ControlFormat.SmallChange.md)
- [Value](Excel.ControlFormat.Value.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]