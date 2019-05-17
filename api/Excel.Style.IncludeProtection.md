---
title: Style.IncludeProtection property (Excel)
keywords: vbaxl10.chm177085
f1_keywords:
- vbaxl10.chm177085
ms.prod: excel
api_name:
- Excel.Style.IncludeProtection
ms.assetid: 666afea1-4a2a-7f44-ecff-d9d44098a527
ms.date: 05/16/2019
localization_priority: Normal
---


# Style.IncludeProtection property (Excel)

**True** if the style includes the **FormulaHidden** and **Locked** protection properties. Read/write **Boolean**.


## Syntax

_expression_.**IncludeProtection**

_expression_ A variable that represents a **[Style](Excel.Style.md)** object.


## Example

This example sets the style attached to cell A1 on Sheet1 to include protection format.

```vb
Worksheets("Sheet1").Range("A1").Style.IncludeProtection = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]