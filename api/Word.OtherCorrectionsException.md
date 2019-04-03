---
title: OtherCorrectionsException object (Word)
keywords: vbawd10.chm2529
f1_keywords:
- vbawd10.chm2529
ms.prod: word
api_name:
- Word.OtherCorrectionsException
ms.assetid: f3c92186-0d3a-0585-b545-3a94e27a7d7b
ms.date: 06/08/2017
localization_priority: Normal
---


# OtherCorrectionsException object (Word)

Represents a single AutoCorrect exception. The  **OtherCorrectionsException** object is a member of the **OtherCorrectionsExceptions** collection.


## Remarks

The  **OtherCorrectionsExceptions** collection includes all words that Microsoft Word won't correct automatically. This list corresponds to the list of AutoCorrect exceptions on the **Other Corrections** tab in the **AutoCorrect Exceptions** dialog box.

Use  **OtherCorrectionsExceptions** (Index), where Index is the AutoCorrect exception name or the index number, to return a single **OtherCorrectionsException** object. The following example deletes "WTop" from the list of AutoCorrect exceptions.




```vb
AutoCorrect.OtherCorrectionsExceptions("WTop").Delete
```

The index number represents the position of the AutoCorrect exception in the  **OtherCorrectionsExceptions** collection. The following example displays the name of the first item in the **OtherCorrectionsExceptions** collection.




```vb
MsgBox AutoCorrect.OtherCorrectionsExceptions(1).Name
```

If the value of the  **OtherCorrectionsAutoAdd** property is **True**, words are automatically added to the list of AutoCorrect exceptions. Use the **Add** method to add an item to the **OtherCorrectionsExceptions** collection. The following example adds "TipTop" to the list of AutoCorrect exceptions.




```vb
AutoCorrect.OtherCorrectionsExceptions.Add Name:="TipTop"
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]