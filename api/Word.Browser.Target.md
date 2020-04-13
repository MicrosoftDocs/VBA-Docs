---
title: Browser.Target property (Word)
keywords: vbawd10.chm154009601
f1_keywords:
- vbawd10.chm154009601
ms.prod: word
api_name:
- Word.Browser.Target
ms.assetid: 138a2e3b-29cb-2523-575b-12ad02e00977
ms.date: 06/08/2017
localization_priority: Normal
---


# Browser.Target property (Word)

Returns or sets the document item that the **Previous** and **Next** methods locate. Read/write **WdBrowseTarget**.


## Syntax

_expression_. `Target`

_expression_ Required. A variable that represents a '[Browser](Word.Browser.md)' object.


## Example

This example moves the insertion point to the next comment in the active document.


```vb
With Application.Browser 
 .Target = wdBrowseComment 
 .Next 
End With
```


## See also


[Browser Object](Word.Browser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]