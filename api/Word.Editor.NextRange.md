---
title: Editor.NextRange property (Word)
keywords: vbawd10.chm225575015
f1_keywords:
- vbawd10.chm225575015
ms.prod: word
api_name:
- Word.Editor.NextRange
ms.assetid: 64c34fd4-2ce8-7d86-0981-1266fe0c7d56
ms.date: 06/08/2017
localization_priority: Normal
---


# Editor.NextRange property (Word)

Returns a  **[Range](Word.Range.md)** object that represents the next range for which a user has permissions to modify.


## Syntax

_expression_. `NextRange`

 _expression_ An expression that returns an '[Editor](Word.Editor.md)' object.


## Remarks

You can also use the  **[GoToEditableRange](Word.Range.GoToEditableRange.md)** method of the **Range** object to return the next range for which a user has permission to modify.


## Example

The following example returns the next range for the first editor in the active selection.


```vb
Dim objEditor As Editor 
Dim objRange As Range 
 
Set objEditor = Selection.Editors(1) 
Set objRange = objEditor.NextRange
```


## See also


[Editor Object](Word.Editor.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]