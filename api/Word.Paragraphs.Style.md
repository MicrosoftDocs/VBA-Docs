---
title: Paragraphs.Style property (Word)
keywords: vbawd10.chm156762212
f1_keywords:
- vbawd10.chm156762212
ms.prod: word
api_name:
- Word.Paragraphs.Style
ms.assetid: 28d5c989-6595-51ea-4fa3-8fb7c0e87b79
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.Style property (Word)

Returns or sets the style for the specified paragraphs. Read/write  **Variant**.


## Syntax

_expression_.**Style**

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Remarks

To set this property, specify the local name of the style, an integer, a  **[WdBuiltinStyle](Word.WdBuiltinStyle.md)** constant, or an object that represents the style.

When you return the style for a range that includes more than one style, only the first character or paragraph style is returned.


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]