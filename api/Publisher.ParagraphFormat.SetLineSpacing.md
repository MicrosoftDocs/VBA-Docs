---
title: ParagraphFormat.SetLineSpacing method (Publisher)
keywords: vbapb10.chm5439511
f1_keywords:
- vbapb10.chm5439511
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.SetLineSpacing
ms.assetid: 32e5b233-8415-2373-7423-18b66df3a5ea
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.SetLineSpacing method (Publisher)

Formats the line spacing of specified paragraphs.


## Syntax

_expression_.**SetLineSpacing** (_Rule_, _Spacing_)

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Rule_|Required| **[PbLineSpacingRule](publisher.pblinespacingrule.md)**|The line spacing to use for the specified paragraphs. Can be one of the **PbLineSpacingRule** constants declared in the Microsoft Publisher type library. |
|_Spacing_|Optional| **Variant**|The spacing (in [points](../language/glossary/vbe-glossary.md#point)) for the specified paragraphs.|


## Example

This example sets the line spacing to double.

```vb
Sub SetLineSpacingForSelection() 
 Selection.TextRange.ParagraphFormat.SetLineSpacing _ 
 Rule:=pbLineSpacingDouble, Spacing:=12 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]