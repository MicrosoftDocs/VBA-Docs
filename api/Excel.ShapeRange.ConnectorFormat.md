---
title: ShapeRange.ConnectorFormat property (Excel)
keywords: vbaxl10.chm640102
f1_keywords:
- vbaxl10.chm640102
api_name:
- Excel.ShapeRange.ConnectorFormat
ms.assetid: cc2c9559-a7f5-8e32-1976-c81e400fb9dd
ms.date: 05/14/2019
ms.localizationpriority: medium
---


# ShapeRange.ConnectorFormat property (Excel)

Returns a **[ConnectorFormat](Excel.ConnectorFormat.md)** object that contains connector formatting properties. Applies to **ShapeRange** objects that represent connectors. Read-only.


## Syntax

_expression_.**ConnectorFormat**

_expression_ An expression that returns a **[ShapeRange](Excel.shaperange.md)** object.


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