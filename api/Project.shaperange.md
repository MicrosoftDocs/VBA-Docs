---
title: ShapeRange object (Project)
ms.prod: project-server
ms.assetid: 315031aa-4b8c-424b-26e7-ce15897beb05
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange object (Project)
Represents a shape range, which is a collection of one or more shapes in a report.
 

## Remarks

Project uses the same Office Art infrastructure that other Office applications use, and adapts Office Art to reports, tables, and charts that can use fields in the active project. However, Project does not implement all  **ShapeRange** operations. For example, Project does not support automatic alignment, distribution, grouping, or merging of shapes in a shape range.
 

 
A shape range can contain a single shape or all the shapes in the report. You can include whichever shapes you want to construct a shape range. For example, you can construct a **ShapeRange** collection that contains the first three shapes in a report, all the shapes in a report, or only the triangle shapes.
 

 
Most operations that you can do with a **Shape** object, you can also do with a **ShapeRange** object that contains only one shape. Some operations, when performed on a **ShapeRange** object that contains more than one shape, shapes of different types, or a shape that is not fully supported in Project, can cause an error. For example, if a shape range contains a rectangle and a chart, and you try to set the **Fill** property, the statement fails because a chart does not implement the **Fill** property. In other cases, for example if you use the **Rotation** property on a shape range that contains a chart and a rectangle, Project rotates the rectangle but silently ignores the chart.
 

 

## Examples

You can return a set of shapes that are specified by the index number or by the shape name. Use  `Shapes.Range(index)`, where _index_ is an array of index numbers or names. For example, both of the following statements are valid:
 

 

```vb
Set myRange1 = theReport.Shapes.Range(Array(1, 2))
Set myRange2 = theReport.Shapes.Range(Array("Textbox 1", "Textbox 2"))
```

To create a **ShapeRange** object that contains all of the shapes in the report, use a statement such as the following:
 

 



```vb
Set allShapes = theReport.Shapes.Range(Array(1, theReport.Shapes.Count))
```

To create a **ShapeRange** object with a single member of the **Shapes** collection, you can use statements such as the following:
 

 



```vb
Set myRange3 = theReport.Shapes.Range(2)
Set myRange4 = theReport.Shapes.Range("Rectangle 2")
```

To perform an operation on a single shape in a **ShapeRange** collection, you can use statements such as the following:
 

 



```vb
myRange1(2).Fill.ForeColor.RGB = RGB(120, 120, 80)
myRange1("Textbox 2").Fill.ForeColor.RGB = RGB(120, 120, 80)
```

Alternately, you can perform an operation directly on a **Shape** object, without using a shape range.
 

 



```vb
theReport.Shapes("Big rectangle").Fill.ForeColor.RGB = RGB(120, 120, 80)
```


## Methods



|**Description**|
|:-----|
|The **Align** method is not implemented in Project.|
|Applies formatting to a shape range, where the formatting information has been copied by using the **[PickUp](Project.shape.pickup.md)** method.|
|Copies the shape range to the Clipboard.|
|Cuts the shape range to the Clipboard.|
|Deletes the shape range.|
|The **Distribute** method is not implemented in Project.|
|Duplicates a shape range and returns a reference to the copy.|
|Flips each shape in the shape range around its horizontal or vertical axis.|
|The **Group** method is not implemented in Project.|
|Moves each shape in the shape range horizontally by the specified number of points.|
|Rotates each shape in the shape range around the z-axis by the specified number of degrees.|
|Moves each shape in the shape range vertically by the specified number of points.|
|Gets an individual  **Shape** object in the shape range collection.|
|The **MergeShapes** method is not implemented in Project.|
|Copies the formatting of the shape range.|
|The **Regroup** method is not implemented in Project.|
|The **RerouteConnections** method is not implemented in Project.|
|Scales the height of the range of shapes by a specified factor.|
|Scales the width of the range of shapes by a specified factor.|
|Selects each shape in a shape range.|
|Applies the formatting of a default shape to each shape in the range.|
|The **Ungroup** method is not implemented in Project.|
|Moves the shape range in front of or behind other shapes (that is, changes the position in the z-order).|

## Properties



|Name|
|:-----|
|[Adjustments](Project.shaperange.adjustments.md)|
|[AlternativeText](Project.shaperange.alternativetext.md)|
|[Application](Project.shaperange.application.md)|
|[AutoShapeType](Project.shaperange.autoshapetype.md)|
|[BackgroundStyle](Project.shaperange.backgroundstyle.md)|
|[BlackWhiteMode](Project.shaperange.blackwhitemode.md)|
|[Callout](Project.shaperange.callout.md)|
|[Chart](Project.shaperange.chart.md)|
|[Child](Project.shaperange.child.md)|
|[ConnectionSiteCount](Project.shaperange.connectionsitecount.md)|
|[Connector](Project.shaperange.connector.md)|
|[ConnectorFormat](Project.shaperange.connectorformat.md)|
|[Count](Project.shaperange.count.md)|
|[Fill](Project.shaperange.fill.md)|
|[Glow](Project.shaperange.glow.md)|
|[GroupItems](Project.shaperange.groupitems.md)|
|[HasChart](Project.shaperange.haschart.md)|
|[HasTable](Project.shaperange.hastable.md)|
|[Height](Project.shaperange.height.md)|
|[HorizontalFlip](Project.shaperange.horizontalflip.md)|
|[ID](Project.shaperange.id.md)|
|[Left](Project.shaperange.left.md)|
|[Line](Project.shaperange.line.md)|
|[LockAspectRatio](Project.shaperange.lockaspectratio.md)|
|[Name](Project.shaperange.name.md)|
|[Nodes](Project.shaperange.nodes.md)|
|[Parent](Project.shaperange.parent.md)|
|[ParentGroup](Project.shaperange.parentgroup.md)|
|[Reflection](Project.shaperange.reflection.md)|
|[Rotation](Project.shaperange.rotation.md)|
|[Script](Project.shaperange.script.md)|
|[Shadow](Project.shaperange.shadow.md)|
|[ShapeStyle](Project.shaperange.shapestyle.md)|
|[SoftEdge](Project.shaperange.softedge.md)|
|[Table](Project.shaperange.table.md)|
|[TextEffect](Project.shaperange.texteffect.md)|
|[TextFrame](Project.shaperange.textframe.md)|
|[TextFrame2](Project.shaperange.textframe2.md)|
|[ThreeD](Project.shaperange.threed.md)|
|[Title](Project.shaperange.title.md)|
|[Top](Project.shaperange.top.md)|
|[Type](Project.shaperange.type.md)|
|[Value](Project.shaperange.value.md)|
|[VerticalFlip](Project.shaperange.verticalflip.md)|
|[Vertices](Project.shaperange.vertices.md)|
|[Visible](Project.shaperange.visible.md)|
|[Width](Project.shaperange.width.md)|
|[ZOrderPosition](Project.shaperange.zorderposition.md)|

## See also


 
[Shapes Object](Project.shapes.md)
 
[Shape Object](Project.shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]