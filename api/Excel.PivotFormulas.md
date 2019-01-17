---
title: PivotFormulas object (Excel)
keywords: vbaxl10.chm232072
f1_keywords:
- vbaxl10.chm232072
ms.prod: excel
api_name:
- Excel.PivotFormulas
ms.assetid: 7139a4bd-f103-7190-004f-7f2261a4391f
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotFormulas object (Excel)

Represents the collection of formulas for a PivotTable report. Each formula is represented by a  **[PivotFormula](Excel.PivotFormula.md)** object.


## Remarks

This object and its associated properties and methods aren't available for OLAP data sources because calculated fields and items aren't supported.


## Example

Use the  **[PivotFormulas](Excel.PivotTable.PivotFormulas.md)** property to return the **PivotFormulas** collection. The following example creates a list of PivotTable formulas for the first PivotTable report on the active worksheet.


```vb
For Each pf in ActiveSheet.PivotTables(1).PivotFormulas 
 Cells(r, 1).Value = pf.Formula 
 r = r + 1 
Next
```


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)


