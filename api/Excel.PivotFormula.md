---
title: PivotFormula object (Excel)
keywords: vbaxl10.chm230072
f1_keywords:
- vbaxl10.chm230072
ms.prod: excel
api_name:
- Excel.PivotFormula
ms.assetid: 2955dad6-d686-1a83-ab56-76a00272c7e2
ms.date: 03/30/2019
localization_priority: Normal
---


# PivotFormula object (Excel)

Represents a formula used to calculate results in a PivotTable report.


## Remarks

This object and its associated properties and methods aren't available for OLAP data sources because calculated fields and items aren't supported.


## Example

Use **[PivotFormulas](Excel.PivotTable.PivotFormulas.md)** (_index_), where _index_ is the formula number or string on the left side of the formula, to return the **PivotFormula** object. 

The following example changes the index number for formula one in the first PivotTable report on the first worksheet so that this formula will be solved after formula two.

```vb
Worksheets(1).PivotTables(1).PivotFormulas(1).Index = 2
```

## Methods

- [Delete](Excel.PivotFormula.Delete.md)

## Properties

- [Application](Excel.PivotFormula.Application.md)
- [Creator](Excel.PivotFormula.Creator.md)
- [Formula](Excel.PivotFormula.Formula.md)
- [Index](Excel.PivotFormula.Index.md)
- [Parent](Excel.PivotFormula.Parent.md)
- [StandardFormula](Excel.PivotFormula.StandardFormula.md)
- [Value](Excel.PivotFormula.Value.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]