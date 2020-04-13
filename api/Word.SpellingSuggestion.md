---
title: SpellingSuggestion object (Word)
keywords: vbawd10.chm2475
f1_keywords:
- vbawd10.chm2475
ms.prod: word
api_name:
- Word.SpellingSuggestion
ms.assetid: 39598da5-8c76-41f3-dcb6-1e1162b30f28
ms.date: 06/08/2017
localization_priority: Normal
---


# SpellingSuggestion object (Word)

Represents a single spelling suggestion for a misspelled word. The **SpellingSuggestion** object is a member of the **[SpellingSuggestions](Word.spellingsuggestions.md)** collection. The **SpellingSuggestions** collection includes all the suggestions for a specified word or for the first word in the specified range.


## Remarks

Use  **GetSpellingSuggestions** (Index), where Index is the index number, to return a single **SpellingSuggestion** object. The following example checks to see whether there are any spelling suggestions for the first word in the active document. If there are, the first suggestion is displayed in a message box.


```vb
If ActiveDocument.Words(1).GetSpellingSuggestions.Count <> 0 Then 
 MsgBox _ 
 ActiveDocument.Words(1).GetSpellingSuggestions.Item(1).Name 
EndIf
```

The **Count** property for the **SpellingSuggestions** object returns 0 (zero) if the word is spelled correctly or if there are no suggestions.


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]