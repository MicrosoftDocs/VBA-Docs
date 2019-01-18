---
title: Shape.BringToFront Method (Visio)
keywords: vis_sdr.chm11216100
f1_keywords:
- vis_sdr.chm11216100
ms.prod: visio
api_name:
- Visio.Shape.BringToFront
ms.assetid: 91689605-16b4-eda5-2513-3e04f78fc13e
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.BringToFront Method (Visio)

Brings the shape or selected shapes to the front of the z-order.


## Syntax

 _expression_. `BringToFront`

 _expression_ A variable that represents a [Shape](./Visio.Shape.md) object.


## Return value

Nothing


## Example

The following macro shows how to bring a shape to the front of the z-order on a page.


```vb
Public Sub BringToFront_Example() 
 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 Dim vsoShape3 As Visio.Shape 
 
 'Draw three rectangles. 
 Set vsoShape1 = ActivePage.DrawRectangle(1, 1, 4, 4) 
 vsoShape1.Text = "1" 
 Set vsoShape2 = ActivePage.DrawRectangle(2, 2, 5, 5) 
 vsoShape2.Text = "2" 
 Set vsoShape3 = ActivePage.DrawRectangle(3, 3, 6, 6) 
 vsoShape3.Text = "3" 
 
 'Bring vsoShape1 to the front of the z-order. 
 vsoShape1.BringToFront 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]