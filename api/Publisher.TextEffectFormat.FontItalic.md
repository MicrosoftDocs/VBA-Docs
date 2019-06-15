---
title: TextEffectFormat.FontItalic property (Publisher)
keywords: vbapb10.chm3735810
f1_keywords:
- vbapb10.chm3735810
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat.FontItalic
ms.assetid: 6594e6f7-e29e-a51d-55b8-d02f1fb9f26a
ms.date: 06/15/2019
localization_priority: Normal
---


# TextEffectFormat.FontItalic property (Publisher)

Sets or returns an **[MsoTriState](Office.MsoTriState.md)** constant that represents whether the font for a dropped capital letter or WordArt text effect is italic. Read/write.


## Syntax

_expression_.**FontItalic**

_expression_ A variable that represents a **[TextEffectFormat](Publisher.TextEffectFormat.md)** object.


## Remarks

The **FontItalic** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library.


## Example

This example makes the dropped capital letter in the specified text frame italic. This example assumes that the specified text frame is formatted with a dropped capital letter.

```vb
Sub BoldDropCap() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.DropCap 
 .FontBold = msoTrue 
 .FontColor.RGB = RGB(Red:=150, Green:=50, Blue:=180) 
 .FontItalic = msoTrue 
 .FontName = "Script MT Bold" 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]