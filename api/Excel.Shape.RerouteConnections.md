---
title: Shape.RerouteConnections method (Excel)
keywords: vbaxl10.chm636082
f1_keywords:
- vbaxl10.chm636082
ms.prod: excel
api_name:
- Excel.Shape.RerouteConnections
ms.assetid: 12e6a6aa-1ddb-392d-14c1-9d57de465c66
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.RerouteConnections method (Excel)

This method reroutes all connectors attached to the specified shape; if the specified shape is a connector, it's rerouted.


## Syntax

_expression_.**RerouteConnections**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Remarks

Reroutes connectors so that they take the shortest possible path between the shapes they connect. To do this, the **RerouteConnections** method may detach the ends of a connector and reattach them to different connecting sites on the connected shapes.

If this method is applied to a connector, only that connector will be rerouted. If this method is applied to a connected shape, all connectors to that shape will be rerouted.


## Example

This example adds two rectangles to _myDocument_, connects them with a curved connector, and then reroutes the connector so that it takes the shortest possible path between the two rectangles. 

Note that the **RerouteConnections** method adjusts the size and position of the connector and determines which connecting sites it attaches to, so the values you initially specify for the _ConnectionSite_ arguments used with the **[BeginConnect](Excel.ConnectorFormat.BeginConnect.md)** and **[EndConnect](Excel.ConnectorFormat.EndConnect.md)** methods are irrelevant.

```vb
Set myDocument = Worksheets(1) 
Set s = myDocument.Shapes 
Set firstRect = s.AddShape(msoShapeRectangle, _ 
 100, 50, 200, 100) 
Set secondRect = s.AddShape(msoShapeRectangle, _ 
 300, 300, 200, 100) 
Set newConnector = s.AddConnector(msoConnectorCurve, _ 
 0, 0, 100, 100) 
With newConnector.ConnectorFormat 
 .BeginConnect firstRect, 1 
 .EndConnect secondRect, 1 
End With 
newConnector.RerouteConnections
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]