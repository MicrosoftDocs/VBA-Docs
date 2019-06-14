---
title: ShapeRange.Adjustments property (Publisher)
keywords: vbapb10.chm2293809
f1_keywords:
- vbapb10.chm2293809
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Adjustments
ms.assetid: 819677e0-806d-a5ac-6fce-f7b0525e63ce
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.Adjustments property (Publisher)

Returns an **[Adjustments](Publisher.Adjustments.md)** collection representing all adjustment handles for the specified **Shape** or **ShapeRange** object.


## Syntax

_expression_.**Adjustments**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Remarks

Adjustment handles correspond to Microsoft Publisher shape sliders.


## Example

This example takes the number of adjustments for a given shape range and assigns it to a variable.

```vb
Public Sub Counter() 
 
 Dim intCount as Integer 
 
 ' A Shape must be in the active publication and selected. 
 intCount = Publisher.ActiveDocument.Selection _ 
 .ShapeRange(1).Adjustments.Count 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]