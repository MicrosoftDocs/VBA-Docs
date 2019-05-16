---
title: SpellingOptions.KoreanUseAutoChangeList property (Excel)
keywords: vbaxl10.chm717081
f1_keywords:
- vbaxl10.chm717081
ms.prod: excel
api_name:
- Excel.SpellingOptions.KoreanUseAutoChangeList
ms.assetid: 9ee57b2d-2a13-8055-d543-234134484fc4
ms.date: 05/16/2019
localization_priority: Normal
---


# SpellingOptions.KoreanUseAutoChangeList property (Excel)

When set to **True**, this enables Microsoft Excel to use the auto-change list for Korean words when using the spelling checker. Read/write **Boolean**.


## Syntax

_expression_.**KoreanUseAutoChangeList**

_expression_ A variable that represents a **[SpellingOptions](Excel.SpellingOptions.md)** object.


## Example

In this example, Microsoft Excel checks to see if the spell checking option to auto-change Korean words is on or off and notifies the user accordingly.

```vb
Sub KoreanSpellCheck() 
 
 If Application.SpellingOptions.KoreanUseAutoChangeList = True Then 
 MsgBox "The spell checking feature to auto-change Korean words is on." 
 Else 
 MsgBox "The spell checking feature to auto-change Korean words is off." 
 End If 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]