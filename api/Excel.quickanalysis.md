---
title: QuickAnalysis object (Excel)
keywords: vbaxl10.chm919072
f1_keywords:
- vbaxl10.chm919072
ms.prod: excel
ms.assetid: cff69157-e5d9-aacb-2569-9727c5f83b0e
ms.date: 04/02/2019
localization_priority: Normal
---


# QuickAnalysis object (Excel)

Enables single-click access to data analysis features such as formulas, conditional formatting, sparklines, tables, charts, and PivotTables.


## Example

This sample illustrates how to use the **Hide** method of the **QuickAnalysis** object. In this example, the _1_ argument specifies that, if displayed, the Conditional Formatting and Sparklines callouts are hidden.

```vb
ActiveWorksheet.QuickAnalysis.Hide(1)
```

## Methods

- [Hide](Excel.quickanalysis.hide.md)
- [Show](Excel.quickanalysis.show.md)

## Properties

- [Application](Excel.quickanalysis.application.md)
- [Creator](Excel.quickanalysis.creator.md)
- [Parent](Excel.quickanalysis.parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
