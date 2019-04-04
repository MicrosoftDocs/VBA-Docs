---
title: Application.MathCoprocessorAvailable property (Excel)
keywords: vbaxl10.chm133161
f1_keywords:
- vbaxl10.chm133161
ms.prod: excel
api_name:
- Excel.Application.MathCoprocessorAvailable
ms.assetid: 9424d6e1-f6f7-cc1b-7d20-987c8ed5e5a2
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.MathCoprocessorAvailable property (Excel)

**True** if a math coprocessor is available. Read-only **Boolean**.


## Syntax

_expression_.**MathCoprocessorAvailable**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays a message box if a math coprocessor isn't available.

```vb
If Not Application.MathCoprocessorAvailable Then 
 MsgBox "This macro requires a math coprocessor" 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]