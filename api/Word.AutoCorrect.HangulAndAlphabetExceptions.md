---
title: AutoCorrect.HangulAndAlphabetExceptions property (Word)
keywords: vbawd10.chm155779085
f1_keywords:
- vbawd10.chm155779085
ms.prod: word
api_name:
- Word.AutoCorrect.HangulAndAlphabetExceptions
ms.assetid: afb525ff-be41-c260-5210-f6ef930b8b04
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrect.HangulAndAlphabetExceptions property (Word)

Returns a  **[HangulAndAlphabetExceptions](Word.hangulandalphabetexceptions.md)** collection that represents the list of Hangul and alphabet AutoCorrect exceptions.


## Syntax

_expression_. `HangulAndAlphabetExceptions`

 _expression_ An expression that returns an '[AutoCorrect](Word.AutoCorrect.md)' object.


## Remarks

This list corresponds to the list of Hangul and alphabet AutoCorrect exceptions on the **Korean** tab in the **AutoCorrect Exceptions** dialog box.

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example prompts the user to delete or keep each hangul and alphabet AutoCorrect exception on the Korean tab in the AutoCorrect Exceptions dialog box.


```vb
For Each anEntry In _ 
 AutoCorrect.HangulAndAlphabetExceptions 
 response = MsgBox("Delete entry: " _ 
 & anEntry.Name, vbYesNoCancel) 
 If response = vbYes Then 
 anEntry.Delete 
 Else 
 If response = vbCancel Then End 
 End If 
Next anEntry
```


## See also


[AutoCorrect Object](Word.AutoCorrect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]