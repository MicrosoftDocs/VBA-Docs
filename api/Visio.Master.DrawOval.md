---
title: Master.DrawOval method (Visio)
keywords: vis_sdr.chm10716210
f1_keywords:
- vis_sdr.chm10716210
ms.prod: visio
api_name:
- Visio.Master.DrawOval
ms.assetid: 092a59d6-1b43-c094-e2ae-480ee7b32b73
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.DrawOval method (Visio)

Adds an oval (ellipse) to the  **Shapes** collection of a master.


## Syntax

_expression_. `DrawOval`( `_x1_` , `_y1_` , `_x2_` , `_y2_` )

_expression_ A variable that represents a **[Master](Visio.Master.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _x1_|Required| **Double**|The x-coordinate of one corner of the ellipse's width-height box.|
| _y1_|Required| **Double**|The y-coordinate of one corner of the ellipse's width-height box.|
| _x2_|Required| **Double**|The x-coordinate of the other corner of the ellipse's width-height box.|
| _y2_|Required| **Double**|The y-coordinate of the other corner of the ellipse's width-height box.|

## Return value

Shape


## Remarks

Using the  **DrawOval** method is equivalent to using the **Ellipse** tool in the application. The arguments are in internal drawing units with respect to the coordinate space of the page, master, or group where the ellipse is being placed.


## Example

The following example shows how to draw an oval (ellipse) on the active page.


```vb
 
Public Sub DrawOval_Example() 
 
 Dim vsoShape As Visio.Shape 
 
 Set vsoShape = ActivePage.DrawOval(1.5, 10.5, 7.5, 6.5) 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]