---
title: Shape.ZOrder method (Project)
ms.prod: project-server
ms.assetid: e8badff9-fbe5-b6b8-8c33-68cfde3bef38
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.ZOrder method (Project)
Moves the shape in front of or behind other shapes (that is, changes the position in the z-order).

## Syntax

_expression_. `ZOrder` _(ZOrderCmd)_

_expression_ A variable that represents a **[Shape](Project.Shape.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ZOrderCmd_|Required|**[MsoZOrderCmd](https://msdn.microsoft.com/library/office/ff861432%28v=office.15%29)**|Specifies where to move the shape relative to the other shapes.|
| _ZOrderCmd_|Required|MSOZORDERCMD||

## Return value

 **Nothing**


## Remarks

Use the **ZOrderPosition** property to determine the current position of a shape in the z-order.


## See also


[Shape Object](Project.shape.md)
[MsoZOrderCmd](https://msdn.microsoft.com/library/office/ff861432%28v=office.15%29)
[ZOrderPosition Property](Project.shaperange.zorderposition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]