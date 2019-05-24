---
title: Workbook.Activate method (Excel)
keywords: vbaxl10.chm199074
f1_keywords:
- vbaxl10.chm199074
ms.prod: excel
api_name:
- Excel.Workbook.Activate
ms.assetid: 628e06b3-ca3f-28cb-e0fd-e696842f69f5
ms.date: 05/25/2019
localization_priority: Normal
---


# Workbook.Activate method (Excel)

Activates the first window associated with the workbook.


## Syntax

_expression_.**Activate**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

This method won't run any Auto_Activate or Auto_Deactivate macros that might be attached to the workbook (use the **[RunAutoMacros](Excel.Workbook.RunAutoMacros.md)** method to run those macros).


## Example

This example activates Book4.xls. If Book4.xls has multiple windows, the example activates the first window, Book4.xls:1.

```vb
Workbooks("BOOK4.XLS").Activate
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
