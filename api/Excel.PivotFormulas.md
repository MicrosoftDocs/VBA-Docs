---
title: PivotFormulas object (Excel)
keywords: vbaxl10.chm232072
f1_keywords:
- vbaxl10.chm232072
ms.prod: excel
api_name:
- Excel.PivotFormulas
ms.assetid: 7139a4bd-f103-7190-004f-7f2261a4391f
ms.date: 03/30/2019
localization_priority: Normal
---


# PivotFormulas object (Excel)

Represents the collection of formulas for a PivotTable report. Each formula is represented by a **[PivotFormula](Excel.PivotFormula.md)** object.


## Remarks

This object and its associated properties and methods aren't available for OLAP data sources because calculated fields and items aren't supported.


## Example

Use the **[PivotFormulas](Excel.PivotTable.PivotFormulas.md)** property of the **PivotTable** object to return the **PivotFormulas** collection. 

The following example creates a list of PivotTable formulas for the first PivotTable report on the active worksheet.

```vb
For Each pf in ActiveSheet.PivotTables(1).PivotFormulas 
 Cells(r, 1).Value = pf.Formula 
 r = r + 1 
Next
```

## Methods

- [Add](Excel.PivotFormulas.Add.md)
- [Item](Excel.PivotFormulas.Item.md)

## Properties

- [Application](Excel.PivotFormulas.Application.md)
- [Count](Excel.PivotFormulas.Count.md)
- [Creator](Excel.PivotFormulas.Creator.md)
- [Parent](Excel.PivotFormulas.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]