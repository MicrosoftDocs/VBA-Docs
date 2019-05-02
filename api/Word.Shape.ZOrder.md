---
title: Shape.ZOrder method (Word)
keywords: vbawd10.chm161480728
f1_keywords:
- vbawd10.chm161480728
ms.prod: word
api_name:
- Word.Shape.ZOrder
ms.assetid: b6729719-44b0-a069-0cbe-b694b88ab65a
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.ZOrder method (Word)

Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).


## Syntax

_expression_.**ZOrder** (_ZOrderCmd_)

 _expression_ An expression that returns a **[Shape](Word.Shape.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ZOrderCmd_|Required| **MsoZOrderCmd**|Specifies where to move the specified shape relative to the other shapes.|

## Return value

Nothing


## Remarks

Use the  **[ZOrderPosition](Word.Shape.ZOrderPosition.md)** property to determine a shape's current position in the z-order.


## Example

This example adds an oval to the active document and then places the oval as second from the back in the z-order if there is at least one other shape on the document.


```vb
With ActiveDocument.Shapes.AddShape(Type:=msoShapeOval, Left:=100, _ 
 Top:=100, Width:=100, Height:=300) 
 While .ZOrderPosition > 2 
 .ZOrder msoSendBackward 
 Wend 
End With
```


## See also


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]