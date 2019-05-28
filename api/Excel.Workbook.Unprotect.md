---
title: Workbook.Unprotect method (Excel)
keywords: vbaxl10.chm199157
f1_keywords:
- vbaxl10.chm199157
ms.prod: excel
api_name:
- Excel.Workbook.Unprotect
ms.assetid: 39387902-a8a4-7bf2-44d7-c5bde6725778
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.Unprotect method (Excel)

Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.


## Syntax

_expression_.**Unprotect** (_Password_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Password_|Optional| **Variant**|A string that denotes the case-sensitive password to use to unprotect the sheet or workbook. If the sheet or workbook isn't protected with a password, this argument is ignored.<br/><br/>If you omit this argument for a sheet that's protected with a password, you'll be prompted for the password. If you omit this argument for a workbook that's protected with a password, the method fails.|

## Remarks

If you forget the password, you cannot unprotect the sheet or workbook. It's a good idea to keep a list of your passwords and their corresponding document names in a safe place.


## Example

This example removes protection from the active workbook.

```vb
ActiveWorkbook.Unprotect
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
