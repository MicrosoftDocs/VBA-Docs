---
title: Worksheet.EnableCalculation property (Excel)
keywords: vbaxl10.chm175079
f1_keywords:
- vbaxl10.chm175079
ms.prod: excel
api_name:
- Excel.Worksheet.EnableCalculation
ms.assetid: fc70ae97-b56b-3b57-6f7b-8438b78c424d
ms.date: 06/08/2017
localization_priority: Priority
---


# Worksheet.EnableCalculation property (Excel)

 **True** if Microsoft Excel automatically recalculates the worksheet when necessary. **False** if Excel doesn't recalculate the sheet. Read/write **Boolean**.


## Syntax

_expression_. `EnableCalculation`

_expression_ A variable that represents a [Worksheet](./Excel.Worksheet.md) object.


## Remarks

When the value of this property is  **False** , you cannot request a recalculation. When you change the value from **False** to **True** , Excel recalculates the worksheet.


## Example

This example sets Microsoft Excel to not recalculate worksheet one automatically.


```vb
Worksheets(1).EnableCalculation = False
```


## See also


[Worksheet Object](Excel.Worksheet.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]