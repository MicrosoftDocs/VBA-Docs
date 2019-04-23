---
title: TwoInitialCapsExceptions object (Word)
keywords: vbawd10.chm2372
f1_keywords:
- vbawd10.chm2372
ms.prod: word
ms.assetid: 21af2d69-8d76-026d-2002-8d69b4ab8aef
ms.date: 06/08/2017
localization_priority: Normal
---


# TwoInitialCapsExceptions object (Word)

A collection of  **[TwoInitialCapsException](Word.TwoInitialCapsException.md)** objects that represent all the items listed in the **Don't correct** box on the **INitial CAps** tab in the **AutoCorrect Exceptions** dialog box.


## Remarks

Use the  **TwoInitialCapsExceptions** property to return the **TwoInitialCapsExceptions** collection. The following example displays the items in this collection.


```vb
For Each aCap In AutoCorrect.TwoInitialCapsExceptions 
 MsgBox aCap.Name 
Next aCap
```

If the  **TwoInitialCapsAutoAdd** property is **True**, words are automatically added to the list of initial-capital exceptions. Use the **Add** method to add an item to the **TwoInitialCapsExceptions** collection. The following example adds "Industry" to the list of initial-capital exceptions.




```vb
AutoCorrect.TwoInitialCapsExceptions.Add Name:="INdustry"
```

Use  **TwoInitialCapsExceptions** (Index), where Index is the initial cap name or the index number, to return a single **TwoInitialCapsException** object. The following example deletes the initial-capital item named "KMenu."




```vb
AutoCorrect.TwoInitialCapsExceptions("KMenu").Delete
```

The index number represents the position of the initial-capital exception in the  **TwoInitialCapsExceptions** collection. The following example displays the name of the first item in the **TwoInitialCapsExceptions** collection.




```vb
MsgBox AutoCorrect.TwoInitialCapsExceptions(1).Name
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]