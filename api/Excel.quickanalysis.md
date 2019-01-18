---
title: QuickAnalysis object (Excel)
keywords: vbaxl10.chm919072
f1_keywords:
- vbaxl10.chm919072
ms.prod: excel
ms.assetid: cff69157-e5d9-aacb-2569-9727c5f83b0e
ms.date: 06/08/2017
localization_priority: Normal
---


# QuickAnalysis object (Excel)

Object that enables single-click access to data analysis features such as formulas, conditional formatting, sparklines, tables, charts, and PivotTables.


## Example

This sample illustrates using the [QuickAnalysis.Hide method (Excel)](Excel.quickanalysis.hide.md) method of the **QuickAnalysis** object. In this example, the _1_ argument specifies that, if displayed, the Conditional Formatting & Sparklines callouts are hidden.


```vb
ActiveWorksheet.QuickAnalysis.Hide(1)
```


## See also

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]