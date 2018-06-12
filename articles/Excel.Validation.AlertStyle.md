---
title: Validation.AlertStyle Property (Excel)
keywords: vbaxl10.chm532074
f1_keywords:
- vbaxl10.chm532074
ms.prod: excel
api_name:
- Excel.Validation.AlertStyle
ms.assetid: de844f58-be45-c4a6-af49-67f669abb626
ms.date: 06/08/2017
---


# Validation.AlertStyle Property (Excel)

Returns the validation alert style. Read-only  **[XlDVAlertStyle](Excel.XlDVAlertStyle.md)** .


## Syntax

 _expression_ . **AlertStyle**

 _expression_ A variable that represents a **Validation** object.


## Remarks

Use the  **[Add](Excel.Validation.Add.md)** method to set the alert style for a range. If the range already has data validation, use the **[Modify](Excel.Validation.Modify.md)** method to change the alert style.


## Example

This example displays the alert style for cell E5.


```vb
MsgBox Range("e5").Validation.AlertStyle
```


## See also


#### Concepts


[Validation Object](Excel.Validation.md)

