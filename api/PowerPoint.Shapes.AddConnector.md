---
title: Shapes.AddConnector method (PowerPoint)
keywords: vbapp10.chm543006
f1_keywords:
- vbapp10.chm543006
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddConnector
ms.assetid: 407eee86-11c1-7bee-ed25-aba71a930a1c
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddConnector method (PowerPoint)

Creates a connector. Returns a **[Shape](PowerPoint.Shape.md)** object that represents the new connector. When a connector is added, it is not connected to anything. Use the **[BeginConnect](PowerPoint.ConnectorFormat.BeginConnect.md)** and **[EndConnect](PowerPoint.ConnectorFormat.EndConnect.md)** methods to attach the beginning and end of a connector to other shapes in the document.


## Syntax

_expression_.**AddConnector** (_Type_, _BeginX_, _BeginY_, _EndX_, _EndY_)

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**[MsoConnectorType](Office.MsoConnectorType.md)**|The type of connector.|
| _BeginX_|Required|**Single**|The horizontal position, measured in points, of the connector's starting point relative to the left edge of the slide.|
| _BeginY_|Required|**Single**|The vertical position, measured in points, of the connector's starting point relative to the top edge of the slide.|
| _EndX_|Required|**Single**|The horizontal position, measured in points, of the connector's ending point relative to the left edge of the slide.|
| _EndY_|Required|**Single**|The vertical position, measured in points, of the connector's ending point relative to the top edge of the slide.|

## Return value

Shape


## Remarks

When you attach a connector to a shape, the size and position of the connector are automatically adjusted, if necessary. Therefore, if you are going to attach a connector to other shapes, the position and dimensions you specify when adding the connector are irrelevant.


## Example

This example adds two rectangles to myDocument and connects them with a curved connector. Note that when you attach the connector to the rectangles, the size and position of the connector are automatically adjusted; therefore, the position and dimensions you specify when adding the callout are irrelevant (dimensions must be nonzero).


```vb
Sub NewConnector() 
 
    Dim shpShapes As Shapes 
    Dim shpFirst As Shape 
    Dim shpSecond As Shape 
 
    Set shpShapes = ActivePresentation.Slides(1).Shapes 
    Set shpFirst = shpShapes.AddShape(Type:=msoShapeRectangle, _ 
        Left:=100, Top:=50, Width:=200, Height:=100) 
    Set shpSecond = shpShapes.AddShape(Type:=msoShapeRectangle, _ 
        Left:=300, Top:=300, Width:=200, Height:=100) 
    With shpShapes.AddConnector(Type:=msoConnectorCurve, BeginX:=0, _ 
            BeginY:=0, EndX:=100, EndY:=100).ConnectorFormat 
        .BeginConnect ConnectedShape:=shpFirst, ConnectionSite:=1 
        .EndConnect ConnectedShape:=shpSecond, ConnectionSite:=1 
        .Parent.RerouteConnections 
    End With 
 
End Sub
```


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]