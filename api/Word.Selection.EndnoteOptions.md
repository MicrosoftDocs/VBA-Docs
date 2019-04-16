---
title: Selection.EndnoteOptions property (Word)
keywords: vbawd10.chm158663681
f1_keywords:
- vbawd10.chm158663681
ms.prod: word
api_name:
- Word.Selection.EndnoteOptions
ms.assetid: 23b7263c-7322-3221-6436-ee0c614fa577
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.EndnoteOptions property (Word)

Returns an  **[EndnoteOptions](Word.EndnoteOptions.md)** object that represents the endnotes in a selection.


## Syntax

_expression_. `EndnoteOptions`

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Example

This example sets the starting number for endnotes in the selected text.


```vb
Sub SetEndnoteOptionsRange() 
 With Selection.EndnoteOptions 
 If .StartingNumber <> 1 Then 
 .StartingNumber = 1 
 End If 
 End With 
End Sub
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]