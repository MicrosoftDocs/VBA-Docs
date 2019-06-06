---
title: ColorFormat.SchemeColor property (Publisher)
keywords: vbapb10.chm2555910
f1_keywords:
- vbapb10.chm2555910
ms.prod: publisher
api_name:
- Publisher.ColorFormat.SchemeColor
ms.assetid: 8b02c85c-a976-7b10-c4ea-6f881d702b55
ms.date: 06/06/2019
localization_priority: Normal
---


# ColorFormat.SchemeColor property (Publisher)

Specifies the color of the current color scheme. Read/write.


## Syntax

_expression_.**SchemeColor**

_expression_ A variable that represents a **[ColorFormat](Publisher.ColorFormat.md)** object.


## Return value

**[PbSchemeColorIndex](Publisher.PbSchemeColorIndex.md)**


## Remarks

The **SchemeColor** property value can be one of the **PbSchemeColorIndex** constants declared in the Microsoft Publisher type library.


## Example

The following example sets the color of the text in shape one on page one of the active publication to accent color five in the current color scheme.

```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Font.Color.SchemeColor = pbSchemeColorAccent5

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]