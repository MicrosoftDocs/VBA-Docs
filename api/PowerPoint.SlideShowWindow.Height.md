---
title: SlideShowWindow.Height property (PowerPoint)
keywords: vbapp10.chm507009
f1_keywords:
- vbapp10.chm507009
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowWindow.Height
ms.assetid: 0108ffd6-2a91-4add-526f-9d34523ca3b9
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowWindow.Height property (PowerPoint)

Returns or sets the height of the specified object, in points. Read/write.


## Syntax

_expression_.**Height**

_expression_ A variable that represents a [SlideShowWindow](PowerPoint.SlideShowWindow.md) object.


## Return value

Single


## Remarks

The **Height** property of a **Shape** object returns or sets the height of the forward-facing surface of the specified shape. This measurement doesn't include shadows or 3D effects.


## Example

This example sets the height of document window two to half the height of the application window.


```vb
Windows(2).Height = Application.Height / 2
```

This example sets the height for row two in the specified table to 100 points (72 points per inch).




```vb
ActivePresentation.Slides(2).Shapes(5).Table.Rows(2).Height = 100
```


## See also


[SlideShowWindow Object](PowerPoint.SlideShowWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]