---
title: TabStop.Alignment property (Word)
keywords: vbawd10.chm156500068
f1_keywords:
- vbawd10.chm156500068
ms.prod: word
api_name:
- Word.TabStop.Alignment
ms.assetid: f27f7f39-439d-0cbf-5538-8d3028abddf1
ms.date: 06/08/2017
localization_priority: Normal
---


# TabStop.Alignment property (Word)

Returns or sets a  **WdTabAlignment** constant that represents the alignment for the specified tab stop. Read/write.


## Syntax

 _expression_. `Alignment`

 _expression_ Required. A variable that represents a '[TabStop](Word.TabStop.md)' object.


## Example

This example centers the first tab stop in the first paragraph of the active document.


```vb
Sub CenterTabStop() 
 ActiveDocument.Paragraphs(1).TabStops(1) _ 
 .Alignment = wdAlignTabCenter 
End Sub
```


## See also


[TabStop Object](Word.TabStop.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]