---
title: ConnectorFormat.EndDisconnect method (Excel)
keywords: vbaxl10.chm646076
f1_keywords:
- vbaxl10.chm646076
ms.prod: excel
api_name:
- Excel.ConnectorFormat.EndDisconnect
ms.assetid: 518345b5-c287-6183-93ae-61c5b56fb9a5
ms.date: 04/23/2019
localization_priority: Normal
---


# ConnectorFormat.EndDisconnect method (Excel)

Detaches the end of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector; the end of the connector remains positioned at a connection site but is no longer connected. 

Use the **[BeginDisconnect](Excel.ConnectorFormat.BeginDisconnect.md)** method to detach the beginning of the connector from a shape.


## Syntax

_expression_.**EndDisconnect**

_expression_ A variable that represents a **[ConnectorFormat](Excel.ConnectorFormat.md)** object.


## Example

This example adds two rectangles to _myDocument_, attaches them with a connector, automatically reroutes the connector along the shortest path, and then detaches the connector from the rectangles.

```vb
Set myDocument = Worksheets(1) 
Set s = myDocument.Shapes 
Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100) 
Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100) 
set c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0) 
with c.ConnectorFormat 
 .BeginConnect firstRect, 1 
 .EndConnect secondRect, 1 
 c.RerouteConnections 
 .BeginDisconnect 
 .EndDisconnect 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]