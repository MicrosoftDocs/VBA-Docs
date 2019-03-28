---
title: ConnectorFormat object (Excel)
keywords: vbaxl10.chm645072
f1_keywords:
- vbaxl10.chm645072
ms.prod: excel
api_name:
- Excel.ConnectorFormat
ms.assetid: 56c97d73-bde2-52ae-2bc3-724d21fdd515
ms.date: 03/29/2019
localization_priority: Normal
---


# ConnectorFormat object (Excel)

Contains properties and methods that apply to connectors.


## Remarks

A connector is a line that attaches two other shapes at points called connection sites. If you rearrange shapes that are connected, the geometry of the connector will be automatically adjusted so that the shapes remain connected.

Connection sites are generally numbered according to the rules presented in the following table.

|Shape type|Connection site numbering scheme|
|:-----|:-----|
|AutoShapes, WordArt, pictures, and OLE objects|The connection sites are numbered starting at the top and proceeding counterclockwise.|
|Freeforms|The connection sites are the vertices, and they correspond to the vertex numbers.|

Use the **[ConnectorFormat](excel.shape.connectorformat.md)** property of the **Shape** object to return a **ConnectorFormat** object. Use the **BeginConnect** and **EndConnect** methods to attach the ends of the connector to other shapes in the document. Use the **[RerouteConnections](Excel.Shape.RerouteConnections.md)** method of the **Shape** object to automatically find the shortest path between the two shapes connected by the connector. Use the **[Connector](Excel.Shape.Connector.md)** property to see whether a shape is a connector.

> [!NOTE] 
> You assign a size and a position when you add a connector to the **[Shapes](excel.shapes.md)** collection, but the size and position are automatically adjusted when you attach the beginning and end of the connector to other shapes in the collection. Therefore, if you intend to attach a connector to other shapes, the initial size and position that you specify are irrelevant.
> 
> Likewise, you specify which connection sites on a shape to attach the connector to when you attach the connector, but using the **RerouteConnections** method after the connector is attached may change which connection sites the connector attaches to, making your original choice of connection sites irrelevant.


## Example

To figure out which number corresponds to which connection site on a complex shape, you can experiment with the shape while the macro recorder is turned on and then examine the recorded code; or you can create a shape, select it, and then run the following example. This code will number each connection site and attach a connector to it.

```vb
Set mainshape = ActiveWindow.Selection.ShapeRange(1) 
With mainshape 
 bx = .Left + .Width + 50 
 by = .Top + .Height + 50 
End With 
With ActiveSheet 
 For j = 1 To mainshape.ConnectionSiteCount 
 With .Shapes.AddConnector(msoConnectorStraight, _ 
 bx, by, bx + 50, by + 50) 
 .ConnectorFormat.EndConnect mainshape, j 
 .ConnectorFormat.Type = msoConnectorElbow 
 .Line.ForeColor.RGB = RGB(255, 0, 0) 
 l = .Left 
 t = .Top 
 End With 
 With .Shapes.AddTextbox(msoTextOrientationHorizontal, _ 
 l, t, 36, 14) 
 .Fill.Visible = False 
 .Line.Visible = False 
 .TextFrame.Characters.Text = j 
 End With 
 Next j 
End With
```

<br/>

The following example adds two rectangles to _myDocument_ and connects them with a curved connector.

```vb
Set myDocument = Worksheets(1) 
Set s = myDocument.Shapes 
Set firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100) 
Set secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100) 
Set c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0) 
With c.ConnectorFormat 
 .BeginConnect ConnectedShape:=firstRect, ConnectionSite:=1 
 .EndConnect ConnectedShape:=secondRect, ConnectionSite:=1 
 c.RerouteConnections 
End With
```


## Methods

- [BeginConnect](Excel.ConnectorFormat.BeginConnect.md)
- [BeginDisconnect](Excel.ConnectorFormat.BeginDisconnect.md)
- [EndConnect](Excel.ConnectorFormat.EndConnect.md)
- [EndDisconnect](Excel.ConnectorFormat.EndDisconnect.md)

## Properties

- [Application](Excel.ConnectorFormat.Application.md)
- [BeginConnected](Excel.ConnectorFormat.BeginConnected.md)
- [BeginConnectedShape](Excel.ConnectorFormat.BeginConnectedShape.md)
- [BeginConnectionSite](Excel.ConnectorFormat.BeginConnectionSite.md)
- [Creator](Excel.ConnectorFormat.Creator.md)
- [EndConnected](Excel.ConnectorFormat.EndConnected.md)
- [EndConnectedShape](Excel.ConnectorFormat.EndConnectedShape.md)
- [EndConnectionSite](Excel.ConnectorFormat.EndConnectionSite.md)
- [Parent](Excel.ConnectorFormat.Parent.md)
- [Type](Excel.ConnectorFormat.Type.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]