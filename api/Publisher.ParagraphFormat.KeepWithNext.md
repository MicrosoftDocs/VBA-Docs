---
title: ParagraphFormat.KeepWithNext property (Publisher)
keywords: vbapb10.chm5439538
f1_keywords:
- vbapb10.chm5439538
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.KeepWithNext
ms.assetid: fb49169d-4718-8ee6-6468-b7cbc8b8a774
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.KeepWithNext property (Publisher)

Sets or returns an **[MsoTriState](office.msotristate.md)** constant that indicates whether the following paragraph will remain in the same text box as the specified paragraph. Read/write.


## Syntax

_expression_.**KeepWithNext**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

MsoTriState


## Remarks

The purpose of **KeepWithNext** is to prevent hanging headings in a document. To do so, you may set this property to **msoTrue** for all headings.

The default setting for this property is **msoFalse**.


## Example

This example sets the **KeepWithNext** property to **msoTrue** for the specified **ParagraphFormat** object.

```vb
Dim objParaForm As ParagraphFormat 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Paragraphs(1).ParagraphFormat 
objParaForm.KeepWithNext = msoTrue
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]