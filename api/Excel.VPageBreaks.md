---
title: VPageBreaks object (Excel)
keywords: vbaxl10.chm166072
f1_keywords:
- vbaxl10.chm166072
ms.prod: excel
api_name:
- Excel.VPageBreaks
ms.assetid: ab8f288a-5235-76c9-7b27-81e542cdd141
ms.date: 04/03/2019
localization_priority: Normal
---


# VPageBreaks object (Excel)

A collection of vertical page breaks within the print area.


## Remarks

Each vertical page break is represented by a **[VPageBreak](Excel.VPageBreak.md)** object.

When the **Application** property, **Count** property, **Creator** property, **Item** property, **Parent** property, or **Add** method is used in conjunction with the **VPageBreaks** property:

- For an automatic print area, the **VPageBreaks** property applies only to the page breaks within the print area.
    
- For a user-defined print area of the same range, the **VPageBreaks** property applies to all of the page breaks.
    

## Example

Use the **[VPageBreaks](Excel.Sheets.VPageBreaks.md)** property of the **Sheets** object to return the **VPageBreaks** collection. Use the **Add** method to add a vertical page break.

If you add a page break that does not intersect the print area, the newly-added **VPageBreak** object does not appear in the **VPageBreaks** collection for the print area. The contents of the collection may change if the print area is resized or redefined.

The following example adds a vertical page break to the left of the active cell.

```vb
ActiveSheet.VPageBreaks.Add Before:=ActiveCell
```

## Methods

- [Add](Excel.VPageBreaks.Add.md)

## Properties

- [Application](Excel.VPageBreaks.Application.md)
- [Count](Excel.VPageBreaks.Count.md)
- [Creator](Excel.VPageBreaks.Creator.md)
- [Item](Excel.VPageBreaks.Item.md)
- [Parent](Excel.VPageBreaks.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]