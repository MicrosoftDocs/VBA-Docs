---
title: HangulAndAlphabetExceptions object (Word)
ms.prod: word
ms.assetid: ddb128f0-3752-5d38-e65a-767f17d86294
ms.date: 06/08/2017
localization_priority: Normal
---


# HangulAndAlphabetExceptions object (Word)

A collection of  **HangulAndAlphabetException** objects that represents all Hangul and alphabet AutoCorrect exceptions.


## Remarks

Use the  **HangulAndAlphabetExceptions** property to return the **HangulAndAlphabetExceptions** collection. The following example displays the items in this collection.


```vb
For Each aHan In AutoCorrect.HangulAndAlphabetExceptions 
 MsgBox aHan.Name 
Next aHan
```

If the value of the  **HangulAndAlphabetAutoAdd** property is **True**, words are automatically added to the list of Hangul and alphabet AutoCorrect exceptions. Use the **Add** method to add an item to the **HangulAndAlphabetExceptions** collection. The following example adds "hello" to the list of alphabet AutoCorrect exceptions.




```vb
AutoCorrect.HangulAndAlphabetExceptions.Add Name:="hello"
```

Use  **HangulAndAlphabetExceptions** (Index), where Index is the Hangul or alphabet AutoCorrect exception name or the index number, to return a single **[HangulAndAlphabetException](Word.HangulAndAlphabetException.md)** object. The following example deletes the alphabet AutoCorrect exception named "goodbye."




```vb
AutoCorrect.HangulAndAlphabetExceptions("goodbye").Delete
```

The index number represents the position of the hangul or alphabet AutoCorrect exception in the  **HangulAndAlphabetExceptions** collection. The following example displays the name of the first item in the **HangulAndAlphabetExceptions** collection.




```vb
MsgBox AutoCorrect.HangulAndAlphabetExceptions(1).Name
```


> [!NOTE] 
> The list of Hangul and alphabet AutoCorrect exceptions corresponds to the list of AutoCorrect exceptions on the  **Korean** tab in the **AutoCorrect Exceptions** dialog box.


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]