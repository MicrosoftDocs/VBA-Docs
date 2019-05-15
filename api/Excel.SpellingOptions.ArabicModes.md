---
title: SpellingOptions.ArabicModes property (Excel)
keywords: vbaxl10.chm717084
f1_keywords:
- vbaxl10.chm717084
ms.prod: excel
api_name:
- Excel.SpellingOptions.ArabicModes
ms.assetid: 0b4fb37e-e5f4-318b-27c1-a90adf39938e
ms.date: 05/16/2019
localization_priority: Normal
---


# SpellingOptions.ArabicModes property (Excel)

Returns or sets the mode for the Arabic spelling checker. Read/write **[XlArabicModes](Excel.XlArabicModes.md)**.


## Syntax

_expression_.**ArabicModes**

_expression_ A variable that represents a **[SpellingOptions](Excel.SpellingOptions.md)** object.


## Example

In this example, Microsoft Excel checks the setting for the spell checking option for Arabic mode and sets it to check for words ending with the letter yaa and words beginning with an alef hamza, if the Arabic mode is not set to this already. Before running this code example, the Arabic modes option must be enabled in the spelling options.

```vb
Sub SpellCheck() 
 
 If Application.SpellingOptions.ArabicModes <> xlArabicBothStrict Then 
 Application.SpellingOptions.ArabicModes = xlArabicBothStrict 
 MsgBox "Spell checking for Arabic mode has been changed to a strict setting." 
 Else 
 MsgBox "Spell checking for Arabic mode is already in a strict setting." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]