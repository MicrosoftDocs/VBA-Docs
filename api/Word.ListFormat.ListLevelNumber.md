---
title: ListFormat.ListLevelNumber property (Word)
keywords: vbawd10.chm163577924
f1_keywords:
- vbawd10.chm163577924
ms.prod: word
api_name:
- Word.ListFormat.ListLevelNumber
ms.assetid: 004c1823-56dd-7a7c-2b0c-8654f0313465
ms.date: 06/08/2017
localization_priority: Normal
---


# ListFormat.ListLevelNumber property (Word)

Returns or sets the list level for the first paragraph in the specified  **ListFormat** object. Read/write **Long**.


## Syntax

 _expression_. `ListLevelNumber`

 _expression_ Required. A variable that represents a '[ListFormat](Word.ListFormat.md)' object.


## Example

This example returns the list level for the third paragraph in the active document.


```vb
lev = ActiveDocument.Paragraphs(3).Range.ListFormat.ListLevelNumber
```


## See also


[ListFormat Object](Word.ListFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]