---
title: Selection.Find property (Word)
keywords: vbawd10.chm158662918
f1_keywords:
- vbawd10.chm158662918
ms.prod: word
api_name:
- Word.Selection.Find
ms.assetid: 66004412-4da2-586d-887c-6f9867e06ea6
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Find property (Word)

Returns a **[Find](Word.Find.md)** object that contains the criteria for a find operation. Read-only.


## Syntax

_expression_.**Find**

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

The selection is changed if the find operation is successful.


## Example

The following example searches forward through the document for the word "Microsoft." If the word is found, it is automatically selected.

```vb
With Selection.Find 
 .Forward = True 
 .ClearFormatting 
 .MatchWholeWord = True 
 .MatchCase = False 
 .Wrap = wdFindContinue 
 .Execute FindText:="Microsoft" 
End With
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
