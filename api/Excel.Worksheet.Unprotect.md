---
title: Worksheet.Unprotect method (Excel)
keywords: vbaxl10.chm174096
f1_keywords:
- vbaxl10.chm174096
ms.prod: excel
api_name:
- Excel.Worksheet.Unprotect
ms.assetid: f955872b-d6bf-5c94-d956-0e84fc7bb9aa
ms.date: 06/08/2017
localization_priority: Priority
---


# Worksheet.Unprotect method (Excel)

Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.


## Syntax

_expression_. `Unprotect`( `_Password_` )

_expression_ A variable that represents a [Worksheet](./Excel.Worksheet.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Password_|Optional| **Variant**|A string that denotes the case-sensitive password to use to unprotect the sheet or workbook. If the sheet or workbook isn't protected with a password, this argument is ignored. If you omit this argument for a sheet that's protected with a password, you'll be prompted for the password. If you omit this argument for a workbook that's protected with a password, the method fails.|

## Remarks

If you forget the password, you cannot unprotect the sheet or workbook. It's a good idea to keep a list of your passwords and their corresponding document names in a safe place.


## Example

This example removes protection from the active workbook.


```vb
ActiveSheet.Unprotect
```


## See also


[Worksheet Object](Excel.Worksheet.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]