---
title: TextEffectFormat.FontBold property (Publisher)
keywords: vbapb10.chm3735809
f1_keywords:
- vbapb10.chm3735809
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat.FontBold
ms.assetid: ab582a4d-92b7-c2b0-e3c3-045e035f68bb
ms.date: 06/15/2019
localization_priority: Normal
---


# TextEffectFormat.FontBold property (Publisher)

Sets or returns an **[MsoTriState](Office.MsoTriState.md)** constant that represents whether the font for a dropped capital letter or WordArt text effect is bold. Read/write.


## Syntax

_expression_.**FontBold**

_expression_ A variable that represents a **[TextEffectFormat](Publisher.TextEffectFormat.md)** object.


## Remarks

The **FontBold** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library.


## Example

This example applies bold formatting to the dropped capital letter in the specified text frame. This example assumes that the specified text frame is formatted with a dropped capital letter.

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