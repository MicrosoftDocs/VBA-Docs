---
title: AutoCorrect.OtherCorrectionsExceptions property (Word)
keywords: vbawd10.chm155779089
f1_keywords:
- vbawd10.chm155779089
ms.prod: word
api_name:
- Word.AutoCorrect.OtherCorrectionsExceptions
ms.assetid: 6353059f-1a87-85e6-8783-f7836ea214f1
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrect.OtherCorrectionsExceptions property (Word)

Returns an  **[OtherCorrectionsExceptions](Word.othercorrectionsexceptions.md)** collection that represents the list of words that Microsoft Word won't correct automatically.


## Syntax

_expression_. `OtherCorrectionsExceptions`

 _expression_ An expression that returns an '[AutoCorrect](Word.AutoCorrect.md)' object.


## Remarks

This list that this property returns corresponds to the list of AutoCorrect exceptions on the **Other Corrections** tab in the **AutoCorrect Exceptions** dialog box.

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example prompts the user to delete or keep each AutoCorrect exception on the **Other Corrections** tab in the **AutoCorrect Exceptions** dialog box.


```vb
For Each anEntry In _ 
 AutoCorrect.OtherCorrectionsExceptions 
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