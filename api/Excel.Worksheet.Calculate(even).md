---
title: Worksheet.Calculate event (Excel)
keywords: vbaxl10.chm502078
f1_keywords:
- vbaxl10.chm502078
ms.prod: excel
api_name:
- Excel.Worksheet.Calculate
ms.assetid: c54b75d0-79dd-3e14-0669-447e740e134b
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Calculate event (Excel)

Occurs after the worksheet is recalculated for the **Worksheet** object.


## Syntax

_expression_.**Calculate**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Return value

**Nothing**


## Example

This example adjusts the size of columns A through F whenever the worksheet is recalculated.

```vb
Private Sub Worksheet_Calculate() 
 Columns("A:F").AutoFit 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
