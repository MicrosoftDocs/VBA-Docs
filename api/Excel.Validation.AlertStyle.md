---
title: Validation.AlertStyle property (Excel)
keywords: vbaxl10.chm532074
f1_keywords:
- vbaxl10.chm532074
ms.prod: excel
api_name:
- Excel.Validation.AlertStyle
ms.assetid: de844f58-be45-c4a6-af49-67f669abb626
ms.date: 05/18/2019
localization_priority: Normal
---


# Validation.AlertStyle property (Excel)

Returns the validation alert style. Read-only **[XlDVAlertStyle](Excel.XlDVAlertStyle.md)**.


## Syntax

_expression_.**AlertStyle**

_expression_ A variable that represents a **[Validation](Excel.Validation.md)** object.


## Remarks

Use the **[Add](Excel.Validation.Add.md)** method to set the alert style for a range. If the range already has data validation, use the **[Modify](Excel.Validation.Modify.md)** method to change the alert style.


## Example

This example displays the alert style for cell E5.

```vb
MsgBox Range("e5").Validation.AlertStyle
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]