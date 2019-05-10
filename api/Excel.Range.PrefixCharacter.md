---
title: Range.PrefixCharacter property (Excel)
keywords: vbaxl10.chm144179
f1_keywords:
- vbaxl10.chm144179
ms.prod: excel
api_name:
- Excel.Range.PrefixCharacter
ms.assetid: 1f7d5fbc-136a-5164-4cec-0054f8bcd0b1
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.PrefixCharacter property (Excel)

Returns the prefix character for the cell. Read-only **Variant**.


## Syntax

_expression_.**PrefixCharacter**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

If the **[TransitionNavigKeys](Excel.Application.TransitionNavigKeys.md)** property is **False**, this prefix character will be `'` for a text label, or blank. 

If the **TransitionNavigKeys** property is **True**, this character will be `'` for a left-justified label, `"` for a right-justified label, `^` for a centered label, `\` for a repeated label, or blank.


## Example

This example displays the prefix character for cell A1 on Sheet1.

```vb
MsgBox "The prefix character is " & _ 
 Worksheets("Sheet1").Range("A1").PrefixCharacter
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]