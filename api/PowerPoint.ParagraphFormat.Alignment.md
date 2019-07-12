---
title: ParagraphFormat.Alignment property (PowerPoint)
keywords: vbapp10.chm576003
f1_keywords:
- vbapp10.chm576003
ms.prod: powerpoint
api_name:
- PowerPoint.ParagraphFormat.Alignment
ms.assetid: 1083d0da-b974-f573-3306-6a865578219b
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.Alignment property (PowerPoint)

Returns or sets the alignment for each paragraph in the specified paragraph format. Read/write.


## Syntax

_expression_.**Alignment**

_expression_ A variable that represents a [ParagraphFormat](PowerPoint.ParagraphFormat.md) object.


## Remarks

The value of the  **Alignment** property can be one of these **PpParagraphAlignment** constants.


||
|:-----|
|**ppAlignCenter**|
|**ppAlignDistribute**|
|**ppAlignJustify**|
|**ppAlignJustifyLow**|
|**ppAlignLeft**|
|**ppAlignmentMixed**|
|**ppAlignRight**|
|**ppAlignThaiDistribute**|

## Example

This example left aligns the paragraphs in shape two on slide one in the active presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes(2) _
    .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
```


## See also


[ParagraphFormat Object](PowerPoint.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]