---
title: Range.ItalicBi property (Word)
keywords: vbawd10.chm157155729
f1_keywords:
- vbawd10.chm157155729
ms.prod: word
api_name:
- Word.Range.ItalicBi
ms.assetid: 69f2ace2-0e12-b704-531c-e4d769d738ec
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.ItalicBi property (Word)

 **True** if the font or range is formatted as italic. Read/write **Long**.


## Syntax

_expression_. `ItalicBi`

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

This property returns  **True**, **False** or **wdUndefined** (for a mixture of italic and non-italic text). Can be set to **True**, **False**, or **wdToggle**.


> [!NOTE] 
> The **ItalicBi** property applies to text in right-to-left languages.


## Example

This example italicizes the first paragraph in the active right-to-left language document.


```vb
ActiveDocument.Paragraphs(1).Range.ItalicBi = True
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]