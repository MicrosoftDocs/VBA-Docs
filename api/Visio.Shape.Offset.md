---
title: Shape.Offset method (Visio)
keywords: vis_sdr.chm11251345
f1_keywords:
- vis_sdr.chm11251345
ms.prod: visio
api_name:
- Visio.Shape.Offset
ms.assetid: 0a82ed87-cc11-77d3-4337-2694f8431a79
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Offset method (Visio)

Offsets a shape a specified amount.


## Syntax

_expression_.**Offset** (_Distance_)

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Distance_|Required| **Double**|Specifies the distance to offset the shape.|

## Return value

Nothing


## Remarks

Calling the  **Offset** method is equivalent to clicking **Offset** in the Microsoft Visio user interface (click **Operations** in the **Shape Design** group on the [Developer](../visio/How-to/run-visio-in-developer-mode.md) tab).

For a specified line or curve, the offset is implemented as a pair of lines or curves that are equidistant from the original line or curve. Offset shapes inherit line patterns from the original shapes. They do not inherit any fill patterns or text from the original shapes.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Offset** method to offset a line shape by a specified amount.


```vb
Public Sub Offset_Example() 
 
 Dim vsoShape As Visio.Shape 
 
 Set vsoShape = Application.ActiveWindow.Page.DrawLine(3, 3, 6, 6) 
 
 ActiveWindow.DeselectAll 
 ActiveWindow.Select vsoShape, visSelect 
 vsoShape.Offset(2) 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]