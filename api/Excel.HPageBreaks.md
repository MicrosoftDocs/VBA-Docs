---
title: HPageBreaks object (Excel)
keywords: vbaxl10.chm163072
f1_keywords:
- vbaxl10.chm163072
ms.prod: excel
api_name:
- Excel.HPageBreaks
ms.assetid: 087106a7-ded7-d672-095d-98e7012fa440
ms.date: 03/30/2019
localization_priority: Normal
---


# HPageBreaks object (Excel)

The collection of horizontal page breaks within the print area.


## Remarks

Each horizontal page break is represented by an **[HPageBreak](Excel.HPageBreak.md)** object.

If you add a page break that does not intersect the print area, the newly-added **HPageBreak** object will not appear in the **HPageBreaks** collection for the print area. The contents of the collection may change if the print area is resized or redefined.

When the **Application** property, **Count** property, **Item** property, **Parent** property, or **Add** method is used in conjunction with the **[HPageBreaks](Excel.Worksheet.HPageBreaks.md)** property of the **Worksheet** object:

- For an automatic print area, the **HPageBreaks** property applies only to the page breaks within the print area.
    
- For a user-defined print area of the same range, the **HPageBreaks** property applies to all of the page breaks.
    
> [!NOTE] 
> There is a limit of 1026 horizontal page breaks per sheet.


## Example

Use the **HPageBreaks** property to return the **HPageBreaks** collection. Use the **Add** method to add a horizontal page break. The following example adds a horizontal page break above the active cell.

```vb
ActiveSheet.HPageBreaks.Add Before:=ActiveCell
```

## Methods

- [Add](Excel.HPageBreaks.Add.md)

## Properties

- [Application](Excel.HPageBreaks.Application.md)
- [Count](Excel.HPageBreaks.Count.md)
- [Creator](Excel.HPageBreaks.Creator.md)
- [Item](Excel.HPageBreaks.Item.md)
- [Parent](Excel.HPageBreaks.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]