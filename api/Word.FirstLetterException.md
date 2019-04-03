---
title: FirstLetterException object (Word)
keywords: vbawd10.chm2373
f1_keywords:
- vbawd10.chm2373
ms.prod: word
api_name:
- Word.FirstLetterException
ms.assetid: e365a683-010a-a074-5563-f0cac1f410b2
ms.date: 06/08/2017
localization_priority: Normal
---


# FirstLetterException object (Word)

Represents an abbreviation excluded from automatic correction. The  **[FirstLetterExceptions](Word.firstletterexceptions.md)** object is a member of the **FirstLetterExceptions** collection.


## Remarks

The  **FirstLetterExceptions** collection includes all the excluded abbreviations.The first character following a period is automatically capitalized when the **CorrectSentenceCaps** property is set to **True**. The character you type following an item in the **FirstLetterExceptions** collection isn't capitalized.

Use  **FirstLetterExceptions** (Index), where Index is the abbreviation or the index number, to return a single **FirstLetterException** object. The following example deletes the abbreviation "appt." from the **[FirstLetterExceptions](Word.firstletterexceptions.md)** collection.




```vb
AutoCorrect.FirstLetterExceptions("appt.").Delete
```

The following example displays the name of the first item in the  **[FirstLetterExceptions](Word.firstletterexceptions.md)** collection.




```vb
MsgBox AutoCorrect.FirstLetterExceptions(1).Name
```

Use the  **Add** method to add an abbreviation to the list of first-letter exceptions. The following example adds the abbreviation "addr." to this list.




```vb
AutoCorrect.FirstLetterExceptions.Add Name:="addr."
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]