---
title: DropCap.FontItalic property (Publisher)
keywords: vbapb10.chm5505030
f1_keywords:
- vbapb10.chm5505030
ms.prod: publisher
api_name:
- Publisher.DropCap.FontItalic
ms.assetid: 57996a71-94db-67b0-ee64-bd79144d01d1
ms.date: 06/07/2019
localization_priority: Normal
---


# DropCap.FontItalic property (Publisher)

Sets or returns an **[MsoTriState](Office.MsoTriState.md)** constant that represents whether the font for a dropped capital letter or WordArt text effect is italic. Read/write.


## Syntax

_expression_.**FontItalic**

_expression_ A variable that represents a **[DropCap](Publisher.DropCap.md)** object.


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