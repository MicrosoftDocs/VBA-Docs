---
title: FirstLetterExceptions object (Word)
ms.prod: word
ms.assetid: 5dc5cc43-a696-d80f-58f9-0f74dfcad0ed
ms.date: 06/08/2017
localization_priority: Normal
---


# FirstLetterExceptions object (Word)

A collection of  **FirstLetterException** objects that represent the abbreviations excluded from automatic correction.


## Remarks

The first character following a period is automatically capitalized when the **CorrectSentenceCaps** property is set to **True**. The **FirstLetterExceptions** collection includes exceptions to this behavior (for example, abbreviations such as "addr." and "apt.").

Use the **FirstLetterExceptions** property to return the **FirstLetterExceptions** collection. The following example deletes the abbreviation "addr." if it is included in the **FirstLetterExceptions** collection.




```vb
For Each aExcept In AutoCorrect.FirstLetterExceptions 
 If aExcept.Name = "addr." Then aExcept.Delete 
Next aExcept
```

The following example creates a new document and inserts all the AutoCorrect first-letter exceptions into it.




```vb
Documents.Add 
For Each aExcept In AutoCorrect.FirstLetterExceptions 
 With Selection 
 .InsertAfter aExcept.Name 
 .InsertParagraphAfter 
 .Collapse Direction:=wdCollapseEnd 
 End With 
Next aExcept
```

Use the **Add** method to add an abbreviation to the list of first-letter exceptions. The following example adds the abbreviation "addr." to this list.




```vb
AutoCorrect.FirstLetterExceptions.Add Name:="addr."
```

Use  **FirstLetterExceptions** (Index), where Index is the abbreviation or the index number, to return a single **[FirstLetterException](Word.FirstLetterException.md)** object. The following example deletes the abbreviation "appt." from the **FirstLetterExceptions** collection.




```vb
AutoCorrect.FirstLetterExceptions("appt.").Delete
```

The following example displays the name of the first item in the **FirstLetterExceptions** collection.




```vb
MsgBox AutoCorrect.FirstLetterExceptions(1).Name
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]