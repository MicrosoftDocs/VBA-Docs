---
title: SpellingOptions.IgnoreMixedDigits property (Excel)
keywords: vbaxl10.chm717077
f1_keywords:
- vbaxl10.chm717077
ms.prod: excel
api_name:
- Excel.SpellingOptions.IgnoreMixedDigits
ms.assetid: 6803fa80-3850-5b34-d22b-3d617c14e537
ms.date: 05/16/2019
localization_priority: Normal
---


# SpellingOptions.IgnoreMixedDigits property (Excel)

**False** instructs Microsoft Excel to check for mixed digits; **True** instructs Excel to ignore mixed digits when checking spelling. Read/write **Boolean**.


## Syntax

_expression_.**IgnoreMixedDigits**

_expression_ A variable that represents a **[SpellingOptions](Excel.SpellingOptions.md)** object.


## Example

In this example, Excel determines what the setting is for checking the spelling of mixed digits and notifies the user.

```vb
Sub SpellingOptionsCheck() 
 
 If Application.SpellingOptions.IgnoreMixedDigits = True Then 
 MsgBox "Spelling options for checking mixed digits is disabled." 
 Else 
 MsgBox "Spelling options for checking mixed digits is enabled." 
 End If 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]