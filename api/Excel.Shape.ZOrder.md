---
title: Shape.ZOrder method (Excel)
keywords: vbaxl10.chm636088
f1_keywords:
- vbaxl10.chm636088
api_name:
- Excel.Shape.ZOrder
ms.assetid: e2eede8f-6e8f-2219-2cb2-47db93e9f90a
ms.date: 05/14/2019
ms.localizationpriority: medium
---


# Shape.ZOrder method (Excel)

Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).


## Syntax

_expression_.**ZOrder** (_ZOrderCmd_)

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ZOrderCmd_|Required| **[MsoZOrderCmd](Office.MsoZOrderCmd.md)**|Specifies where to move the specified shape relative to the other shapes.|

## Remarks

Use the **[ZOrderPosition](Excel.Shape.ZOrderPosition.md)** property to determine a shape's current position in the z-order.


## Example

This example adds an oval to _myDocument_ and then places the oval second from the back in the z-order if there is at least one other shape on the document.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300) 
    While .ZOrderPosition > 2 
        .ZOrder msoSendBackward 
    Wend 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]