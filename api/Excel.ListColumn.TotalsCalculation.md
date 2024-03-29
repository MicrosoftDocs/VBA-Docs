---
title: ListColumn.TotalsCalculation property (Excel)
keywords: vbaxl10.chm738079
f1_keywords:
- vbaxl10.chm738079
api_name:
- Excel.ListColumn.TotalsCalculation
ms.assetid: bb8c1dd1-1ee6-3ef8-8af4-2b3f24eb642d
ms.date: 04/30/2019
ms.localizationpriority: medium
---


# ListColumn.TotalsCalculation property (Excel)

Determines the type of calculation in the Totals row of the list column based on the value of the **[XlTotalsCalculation](Excel.XlTotalsCalculation.md)** enumeration. Read/write.


## Syntax

_expression_.**TotalsCalculation**

_expression_ A variable that represents a **[ListColumn](Excel.ListColumn.md)** object.


## Remarks

The Totals row doesn't need to be showing to set this property. There is no fixed "default" value for this property. Excel may change the state of this property as other columns are added or deleted.


## Example

```vb
ActiveSheet.ListColumns(1).TotalsCalculation=xlTotalsCalculationSum
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]