---
title: Paragraph.BaseLineAlignment property (Word)
keywords: vbawd10.chm156696699
f1_keywords:
- vbawd10.chm156696699
ms.prod: word
api_name:
- Word.Paragraph.BaseLineAlignment
ms.assetid: 27639ce6-4ef1-4252-873d-270ae19daba8
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.BaseLineAlignment property (Word)

Returns or sets a  **WdBaselineAlignment** constant that represents the vertical position of fonts on a line. Read/write.


## Syntax

_expression_. `BaseLineAlignment`

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Example

This example sets Microsoft Word to automatically adjust the baseline font alignment in the active document.


```vb
ActiveDocument.BaseLineAlignment = wdBaselineAlignAuto
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]