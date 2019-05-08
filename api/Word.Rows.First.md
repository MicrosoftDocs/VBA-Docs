---
title: Rows.First property (Word)
keywords: vbawd10.chm155975690
f1_keywords:
- vbawd10.chm155975690
ms.prod: word
api_name:
- Word.Rows.First
ms.assetid: 9e879fdf-bc21-cd19-37e9-bf44c06b3416
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows.First property (Word)

Returns a  **[Row](Word.Row.md)** object that represents the first item in the **Rows** collection.


## Syntax

_expression_. `First`

_expression_ Required. A variable that represents a **[Rows](Word.Rows.md)** object.


## Example

This example applies shading and a bottom border to the first row in the first table of the active document.


```vb
ActiveDocument.Tables(1).Borders.Enable = False 
With ActiveDocument.Tables(1).Rows.First 
 .Shading.Texture = wdTexture10Percent 
 .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle 
End With
```


## See also


[Rows Collection Object](Word.rows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]