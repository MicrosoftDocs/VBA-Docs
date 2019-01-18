---
title: VPageBreaks object (Excel)
keywords: vbaxl10.chm166072
f1_keywords:
- vbaxl10.chm166072
ms.prod: excel
api_name:
- Excel.VPageBreaks
ms.assetid: ab8f288a-5235-76c9-7b27-81e542cdd141
ms.date: 06/08/2017
localization_priority: Normal
---


# VPageBreaks object (Excel)

A collection of vertical page breaks within the print area.


## Remarks

Each vertical page break is represented by a  **[VPageBreak](Excel.VPageBreak.md)** object.

When the [Application](Excel.VPageBreaks.Application.md) property, **[Count](Excel.VPageBreaks.Count.md)** property, **[Creator](Excel.LineFormat.Creator.md)** property, **[Item](Excel.VPageBreaks.Item.md)** property, **[Parent](Excel.VPageBreaks.Parent.md)** property or **[Add](Excel.VPageBreaks.Add.md)** method is used in conjunction with the **VPageBreaks** property:


- For an automatic print area, the  **VPageBreaks** property applies only to the page breaks within the print area.
    
- For a user-defined print area of the same range, the  **VPageBreaks** property applies to all of the page breaks.
    

## Example

Use the  **[VPageBreaks](Excel.Sheets.VPageBreaks.md)** property to return the **VPageBreaks** collection. Use the **[Add](Excel.VPageBreaks.Add.md)** method to add a vertical page break.

If you add a page break that does not intersect the print area, then the newly-added  **VPageBreak** object will not appear in the **VPageBreaks** collection for the print area. The contents of the collection may change if the print area is resized or redefined.

The following example adds a vertical page break to the left of the active cell.




```vb
ActiveSheet.VPageBreaks.Add Before:=ActiveCell
```


## See also


[Excel Object Model Reference](./overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]