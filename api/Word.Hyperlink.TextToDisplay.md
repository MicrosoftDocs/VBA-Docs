---
title: Hyperlink.TextToDisplay property (Word)
keywords: vbawd10.chm161285108
f1_keywords:
- vbawd10.chm161285108
ms.prod: word
api_name:
- Word.Hyperlink.TextToDisplay
ms.assetid: 9b9f73cd-bf4e-367e-c901-746b85da9f9c
ms.date: 06/08/2017
localization_priority: Normal
---


# Hyperlink.TextToDisplay property (Word)

Returns or sets the specified hyperlink's visible text in a document. Read/write  **String**.


## Syntax

 _expression_. `TextToDisplay`

 _expression_ An expression that returns a '[Hyperlink](Word.Hyperlink.md)' object.


## Example

This example sets the display text for the first hyperlink in the active document.


```vb
ActiveDocument.Hyperlinks(1).TextToDisplay = _ 
 "Follow this link for more information..."
```


## See also


[Hyperlink Object](Word.Hyperlink.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]