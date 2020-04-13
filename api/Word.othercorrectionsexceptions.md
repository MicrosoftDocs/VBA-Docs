---
title: OtherCorrectionsExceptions object (Word)
ms.prod: word
ms.assetid: f72135be-9a82-2c45-1835-8cabb18869de
ms.date: 06/08/2017
localization_priority: Normal
---


# OtherCorrectionsExceptions object (Word)

A collection of  **OtherCorrectionsException** objects that represents the list of words that Microsoft Word won't correct automatically.


## Remarks

This list corresponds to the list of AutoCorrect exceptions on the **Other Corrections** tab in the **AutoCorrect Exceptions** dialog box.

Use the **OtherCorrectionsExceptions** property to return the **OtherCorrectionsExceptions** collection. The following example displays the items in this collection.




```vb
For Each aCap In AutoCorrect.OtherCorrectionsExceptions 
 MsgBox aCap.Name 
Next aCap
```

If the value of the **OtherCorrectionsAutoAdd** property is **True**, words are automatically added to the list of AutoCorrect exceptions. Use the **Add** method to add an item to the **OtherCorrectionsExceptions** collection. The following example adds "TipTop" to the list of AutoCorrect exceptions.




```vb
AutoCorrect.OtherCorrectionsExceptions.Add Name:="TipTop"
```

Use  **OtherCorrectionsExceptions** (Index), where Index is the name or the index number, to return a single **OtherCorrectionsException** object. The following example deletes "WTop" from the list of AutoCorrect exceptions.




```vb
AutoCorrect.OtherCorrectionsExceptions("WTop").Delete
```

The index number represents the position of the AutoCorrect exception in the **OtherCorrectionsExceptions** collection. The following example displays the name of the first item in the **OtherCorrectionsExceptions** collection.




```vb
MsgBox AutoCorrect.OtherCorrectionsExceptions(1).Name
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]