---
title: Selection.WholeStory method (Word)
keywords: vbawd10.chm158663180
f1_keywords:
- vbawd10.chm158663180
ms.prod: word
api_name:
- Word.Selection.WholeStory
ms.assetid: ecd50a78-ecbd-75a9-2565-31d7e6ac449a
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.WholeStory method (Word)

Expands a selection to include the entire story.


## Syntax

_expression_. `WholeStory`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

The following instructions, where  _objSel_ is a valid **Selection** object, are functionally equivalent:


```vb
objSel.WholeStory 
objSel.Expand Unit:=wdStory
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
