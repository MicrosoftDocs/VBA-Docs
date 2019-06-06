---
title: CellBorder.Color property (Publisher)
keywords: vbapb10.chm5242882
f1_keywords:
- vbapb10.chm5242882
ms.prod: publisher
api_name:
- Publisher.CellBorder.Color
ms.assetid: 59a43522-f0df-fe1a-6e35-19cb012b103f
ms.date: 06/06/2019
localization_priority: Normal
---


# CellBorder.Color property (Publisher)

Returns a **[ColorFormat](Publisher.ColorFormat.md)** object representing the color information for the specified object.


## Syntax

_expression_.**Color**

_expression_ A variable that represents a **[CellBorder](Publisher.CellBorder.md)** object.


## Example

This example tests the font color of the first story in the active document and tells the user whether the font color is black.

```vb
Sub FontColor() 
 
 If Application.ActiveDocument.Stories(1) _ 
 .TextRange.Font.Color.RGB = RGB(Red:=0, Green:=0, Blue:=0) Then 
 MsgBox "Your font color is black" 
 Else 
 MsgBox "Your font color is not black" 
 End If 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]