---
title: Workbooks.Close method (Excel)
keywords: vbaxl10.chm203074
f1_keywords:
- vbaxl10.chm203074
ms.prod: excel
api_name:
- Excel.Workbooks.Close
ms.assetid: d798166c-6b27-16a1-0b64-8f547978e371
ms.date: 05/18/2019
localization_priority: Normal
---


# Workbooks.Close method (Excel)

Closes the object.


## Syntax

_expression_.**Close**

_expression_ A variable that represents a **[Workbooks](Excel.Workbooks.md)** object.


## Remarks

Closing a workbook from Visual Basic doesn't run any Auto_Close macros in the workbook. Use the **[RunAutoMacros](Excel.Workbook.RunAutoMacros.md)** method to run the auto close macros.


## Example

This example closes all open workbooks. If there are changes in any open workbook, Microsoft Excel displays the appropriate prompts and dialog boxes for saving changes.

```vb
Workbooks.Close
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
