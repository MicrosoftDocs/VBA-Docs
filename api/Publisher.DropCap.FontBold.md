---
title: DropCap.FontBold property (Publisher)
keywords: vbapb10.chm5505029
f1_keywords:
- vbapb10.chm5505029
ms.prod: publisher
api_name:
- Publisher.DropCap.FontBold
ms.assetid: 7e1b9b51-258d-080c-e5ae-cdc9d6a2ba64
ms.date: 06/07/2019
localization_priority: Normal
---


# DropCap.FontBold property (Publisher)

Sets or returns an **[MsoTriState](Office.MsoTriState.md)** constant that represents whether the font for a dropped capital letter or WordArt text effect is bold. Read/write.


## Syntax

_expression_.**FontBold**

_expression_ A variable that represents a **[DropCap](Publisher.DropCap.md)** object.


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