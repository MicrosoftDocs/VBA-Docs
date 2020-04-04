---
title: Shape.ConnectorFormat property (PowerPoint)
keywords: vbapp10.chm547021
f1_keywords:
- vbapp10.chm547021
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.ConnectorFormat
ms.assetid: 6c3f7f40-02a8-73ff-5829-7994ba1495d2
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.ConnectorFormat property (PowerPoint)

Returns a **[ConnectorFormat](PowerPoint.ConnectorFormat.md)** object that contains connector formatting properties. Applies to **Shape** or **ShapeRange** objects that represent connectors. Read-only.


## Syntax

_expression_.**ConnectorFormat**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

ConnectorFormat


## Example

This example adds two rectangles to _myDocument_, attaches them with a connector, automatically reroutes the connector along the shortest path, and then detaches the connector from the rectangles.


```vb
Set myDocument = ActivePresentation.Slides(1)

Set s = myDocument.Shapes

Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)

Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)

With s.AddConnector(msoConnectorCurve, 0, 0, 0, 0).ConnectorFormat

    .BeginConnect firstRect, 1

    .EndConnect secondRect, 1

    .Parent.RerouteConnections

    .BeginDisconnect

    .EndDisconnect

End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]