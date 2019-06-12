---
title: ParagraphFormat.ListIndent property (Publisher)
keywords: vbapb10.chm5439522
f1_keywords:
- vbapb10.chm5439522
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.ListIndent
ms.assetid: b42000ea-0636-88cf-b7ed-c71384a2b0d5
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.ListIndent property (Publisher)

Returns or sets a **Single** that represents the list indent value (in [points](../language/glossary/vbe-glossary.md#point)) for the specified **ParagraphFormat** object. Read/write.


## Syntax

_expression_.**ListIndent**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

Single


## Example

This example sets the **ListIndent** property of a **ParagraphFormat** object to 0.25 inches. The **[InchesToPoints](publisher.application.inchestopoints.md)** method is used to convert inches to points.

```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 .ListIndent = InchesToPoints(0.25) 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]