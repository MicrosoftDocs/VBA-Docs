---
title: ParagraphFormat.KeepLinesTogether property (Publisher)
keywords: vbapb10.chm5439537
f1_keywords:
- vbapb10.chm5439537
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.KeepLinesTogether
ms.assetid: a0f3f2f0-d986-4928-3c4f-0665711a6876
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.KeepLinesTogether property (Publisher)

Sets or returns an **[MsoTriState](office.msotristate.md)** constant that indicates whether all lines in the specified paragraph will remain in the same text box. Read/write.


## Syntax

_expression_.**KeepLinesTogether**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

msoTriState


## Remarks

This option ensures that there is not a text frame or column break between the lines of the specified paragraph. If the paragraphs are too large for the text frame or column, the first line will start at the top of the next text frame or column.

The default setting for this property is **msoFalse**.


## Example

This example sets the **KeepLinesTogether** property to **msoTrue** for the specified **ParagraphFormat** object.

```vb
Dim objParaForm As ParagraphFormat 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Paragraphs(1).ParagraphFormat 
objParaForm.KeepLinesTogether = msoTrue 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]