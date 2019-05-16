---
title: SpellingOptions.SuggestMainOnly property (Excel)
keywords: vbaxl10.chm717076
f1_keywords:
- vbaxl10.chm717076
ms.prod: excel
api_name:
- Excel.SpellingOptions.SuggestMainOnly
ms.assetid: f4a5aa0a-78be-bd98-22e8-b85eac0f4428
ms.date: 05/16/2019
localization_priority: Normal
---


# SpellingOptions.SuggestMainOnly property (Excel)

When set to **True**, instructs Microsoft Excel to suggest words from only the main dictionary when using the spelling checker. **False** removes the limits of suggesting words from only the main dictionary when using the spelling checker. Read/write **Boolean**.


## Syntax

_expression_.**SuggestMainOnly**

_expression_ A variable that represents a **[SpellingOptions](Excel.SpellingOptions.md)** object.


## Example

In this example, Microsoft Excel checks the spelling checking options for suggesting words only from the main dictionary and reports the status to the user.

```vb
Sub UsingMainDictionary() 
 
 ' Check the setting of suggesting words only from the main dictionary. 
 If Application.SpellingOptions.SuggestMainOnly = True Then 
 MsgBox "Spell checking option suggestions will only come from the main dictionary." 
 Else 
 MsgBox "Spell checking option suggestions are not limited to the main dictionary." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]