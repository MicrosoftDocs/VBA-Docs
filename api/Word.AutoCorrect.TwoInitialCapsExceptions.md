---
title: AutoCorrect.TwoInitialCapsExceptions property (Word)
keywords: vbawd10.chm155779081
f1_keywords:
- vbawd10.chm155779081
ms.prod: word
api_name:
- Word.AutoCorrect.TwoInitialCapsExceptions
ms.assetid: c301d210-c583-a092-4840-ac8efed80c86
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrect.TwoInitialCapsExceptions property (Word)

Returns a  **[TwoInitialCapsExceptions](Word.twoinitialcapsexceptions.md)** collection that represents the list of terms containing mixed capitalization that Word won't correct automatically.


## Syntax

_expression_. `TwoInitialCapsExceptions`

 _expression_ An expression that returns an '[AutoCorrect](Word.AutoCorrect.md)' object.


## Remarks

This list corresponds to the list of AutoCorrect exceptions on the INitial CAps tab in the  **AutoCorrect Exceptions** dialog box (**AutoCorrect Options** command, **Tools** menu). For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example prompts the user to delete or keep each AutoCorrect Initial Caps exception.


```vb
For Each anEntry In AutoCorrect.TwoInitialCapsExceptions 
 response = MsgBox ("Delete entry: " _ 
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