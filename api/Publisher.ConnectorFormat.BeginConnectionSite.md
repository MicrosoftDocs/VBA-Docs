---
title: ConnectorFormat.BeginConnectionSite property (Publisher)
keywords: vbapb10.chm3211522
f1_keywords:
- vbapb10.chm3211522
ms.prod: publisher
api_name:
- Publisher.ConnectorFormat.BeginConnectionSite
ms.assetid: 24a9246e-270f-7289-971d-8763acfaf02d
ms.date: 06/06/2019
localization_priority: Normal
---


# ConnectorFormat.BeginConnectionSite property (Publisher)

Returns a **Long** indicating the connection site to which the beginning of a connector is connected. Read-only.


## Syntax

_expression_.**BeginConnectionSite**

_expression_ A variable that represents a **[ConnectorFormat](Publisher.ConnectorFormat.md)** object.


## Return value

Long


## Remarks

If the beginning of the specified connector isn't attached to a shape, this property generates an error.

Use the **[EndConnectionSite](Publisher.ConnectorFormat.EndConnectionSite.md)** property to return the site to which the end of a connector is connected.


## Example

This example assumes that the first page in the active publication already contains two shapes attached by a connector named Conn1To2. The code adds a rectangle and a connector to the first page. The beginning of the new connector will be attached to the same connection site as the beginning of the connector named Conn1To2, and the end of the new connector will be attached to connection site one on the new rectangle.

```vb
Dim shpNew As Shape 
Dim intSite As Integer 
Dim shpOld As Shape 
 
With ActiveDocument.Pages(1).Shapes 
 
 ' Add new rectangle. 
 Set shpNew = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=450, Top:=190, Width:=200, Height:=100) 
 
 ' Add new connector. 
 .AddConnector(Type:=msoConnectorCurve, _ 
 BeginX:=0, BeginY:=0, EndX:=10, EndY:=10) _ 
 .Name = "Conn1To3" 
 
 ' Get connection site number of old shape, and set 
 ' reference to old shape. 
 With .Item("Conn1To2").ConnectorFormat 
 intSite = .BeginConnectionSite 
 Set shpOld = .BeginConnectedShape 
 End With 
 
 ' Connect new connector to old shape and new rectangle. 
 With .Item("Conn1To3").ConnectorFormat 
 .BeginConnect ConnectedShape:=shpOld, _ 
 ConnectionSite:=intSite 
 .EndConnect ConnectedShape:=shpNew, _ 
 ConnectionSite:=1 
 End With 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]