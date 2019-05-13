---
title: Shape.ConnectorFormat property (Excel)
keywords: vbaxl10.chm636095
f1_keywords:
- vbaxl10.chm636095
ms.prod: excel
api_name:
- Excel.Shape.ConnectorFormat
ms.assetid: 4c000a5c-eed2-e93c-e801-999c96750c9e
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.ConnectorFormat property (Excel)

Returns a **[ConnectorFormat](Excel.ConnectorFormat.md)** object that contains connector formatting properties. Applies to a **Shape** object that represents connectors. Read-only.


## Syntax

_expression_.**ConnectorFormat**

_expression_ An expression that returns a **[Shape](Excel.Shape.md)** object.


## Example

This example adds two rectangles to _myDocument_, attaches them with a connector, automatically reroutes the connector along the shortest path, and then detaches the connector from the rectangles.

```vb
Set myDocument = Worksheets(1) 
Set s = myDocument.Shapes 
Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100) 
Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100) 
Set c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0) 
with c.ConnectorFormat 
 .BeginConnect firstRect, 1 
 .EndConnect secondRect, 1 
 c.RerouteConnections 
 .BeginDisconnect 
 .EndDisconnect 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]