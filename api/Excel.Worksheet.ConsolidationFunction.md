---
title: Worksheet.ConsolidationFunction property (Excel)
keywords: vbaxl10.chm175087
f1_keywords:
- vbaxl10.chm175087
ms.prod: excel
api_name:
- Excel.Worksheet.ConsolidationFunction
ms.assetid: 4a356e31-723c-88e9-575b-b5a7c5e67757
ms.date: 06/08/2017
localization_priority: Normal
---


# Worksheet.ConsolidationFunction property (Excel)

Returns the function code used for the current consolidation. Can be one of the constants of  **[xlConsolidationFunction](Excel.XlConsolidationFunction.md)**. Read-only **Long**.


## Syntax

_expression_. `ConsolidationFunction`

_expression_ A variable that represents a [Worksheet](./Excel.Worksheet.md) object.


## Example

This example displays a message box if the current consolidation is using the SUM function.


```vb
If Worksheets("Sheet1").ConsolidationFunction = xlSum Then 
 MsgBox "Sheet1 uses the SUM function for consolidation." 
End If
```


## See also


[Worksheet Object](Excel.Worksheet.md)

