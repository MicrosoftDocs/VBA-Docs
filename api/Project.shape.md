---
title: Shape object (Project)
ms.prod: project-server
ms.assetid: d2b32bcd-5595-a4a7-9772-feb25fd0103a
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape object (Project)
Represents an object in a Project report, such as a chart, report table, text box, freeform drawing, or picture.
 

## Remarks

The  **Shape** object is a member of the **[Shapes](Project.shapes.md)** collection, which includes all of the shapes in the report.
 

 

> [!NOTE] 
> Macro recording for the  **Shape** object is not implemented. That is, when you record a macro in Project and manually add a shape or edit shape elements, the steps for adding and manipulating the shape are not recorded.
 

There are three objects that represent shapes: the  **Shapes** collection, which represents all the shapes on a document; the **ShapeRange** object, which represents a specified subset of the shapes on a document (for example, a **ShapeRange** object could represent shapes one and four on the document, or it could represent all the selected shapes on the document); and the **Shape** object, which represents a single shape on a document. If you want to work with several shapes at the same time or with shapes within the selection, use a **ShapeRange** collection.
 

 
Use  `Shapes(Index)`, where  _Index_ is the shape name or the index number, to return a single **Shape** object.
 

 

## Example

In the following example, the  **TestTextShape** macro creates a textbox shape, adds some text, and changes the shape style, fill, line, shadow, and reflection properties. The **FlipShape** macro flips the shape from top to bottom.
 

 

```vb
Sub TestTextShape()
    Dim theReport As Report
    Dim textShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    
    Set theReport = ActiveProject.Reports(reportName)
    Set textShape = theReport.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 30, 300, 100)
    textShape.Name = "TestTextBox"
    
    textShape.TextFrame2.TextRange.Characters.Text = "This is a test. It is only a test. " _
        & "If it had been real information, there would be some real text here."
    textShape.TextFrame2.TextRange.Characters(1, 15).ParagraphFormat.FirstLineIndent = 0
    
    ' Set the font for the first 15 characters to dark blue bold.
    With textShape.TextFrame2.TextRange.Characters(1, 15).Font
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent5
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Bold = msoTrue
    End With
    
    textShape.ShapeStyle = msoShapeStylePreset42
    
    With textShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 0)
        .Transparency = 0
        '.Solid
    End With
   
    With textShape.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With

    textShape.Shadow.Type = msoShadow22
    textShape.Reflection.Type = msoReflectionType3
End Sub

Sub FlipShape()
    Dim theReport As Report
    Dim theShape As Shape
    Dim reportName As String
    Dim shapeName As String
    
    reportName = "Simple scalar chart"
    shapeName = "TestTextBox"
    
    Set theShape = ActiveProject.Reports(reportName).Shapes(shapeName)

    theShape.Flip msoFlipVertical
    theShape.Select
End Sub
```

Figure 1 shows the result, where the shape is selected to make the ribbon  **FORMAT** tab under **DRAWING TOOLS** available, although the active tab is **DESIGN** under **REPORT TOOLS**. If the shape were not selected,  **DRAWING TOOLS** and the **FORMAT** tab would not be visible.
 

 

**Figure 1. Testing the Shape object model**

 
![Testing the Shape object model](../images/pj15_VBA_ShapeObject.gif)
 

 

## Methods



|Name|
|:-----|
|[Apply](Project.shape.apply.md)|
|[Copy](Project.shape.copy.md)|
|[Cut](Project.shape.cut.md)|
|[Delete](Project.shape.delete.md)|
|[Duplicate](Project.shape.duplicate.md)|
|[Flip](Project.shape.flip.md)|
|[IncrementLeft](Project.shape.incrementleft.md)|
|[IncrementRotation](Project.shape.incrementrotation.md)|
|[IncrementTop](Project.shape.incrementtop.md)|
|[PickUp](Project.shape.pickup.md)|
|[RerouteConnections](Project.shape.rerouteconnections.md)|
|[ScaleHeight](Project.shape.scaleheight.md)|
|[ScaleWidth](Project.shape.scalewidth.md)|
|[Select](Project.shape.select.md)|
|[SetShapesDefaultProperties](Project.shape.setshapesdefaultproperties.md)|
|[Ungroup](Project.shape.ungroup.md)|
|[ZOrder](Project.shape.zorder.md)|

## Properties



|Name|
|:-----|
|[Adjustments](Project.shape.adjustments.md)|
|[AlternativeText](Project.shape.alternativetext.md)|
|[Application](Project.shape.application.md)|
|[AutoShapeType](Project.shape.autoshapetype.md)|
|[BackgroundStyle](Project.shape.backgroundstyle.md)|
|[BlackWhiteMode](Project.shape.blackwhitemode.md)|
|[Callout](Project.shape.callout.md)|
|[Chart](Project.shape.chart.md)|
|[Child](Project.shape.child.md)|
|[ConnectionSiteCount](Project.shape.connectionsitecount.md)|
|[Connector](Project.shape.connector.md)|
|[ConnectorFormat](Project.shape.connectorformat.md)|
|[Fill](Project.shape.fill.md)|
|[Glow](Project.shape.glow.md)|
|[GroupItems](Project.shape.groupitems.md)|
|[HasChart](Project.shape.haschart.md)|
|[HasTable](Project.shape.hastable.md)|
|[Height](Project.shape.height.md)|
|[HorizontalFlip](Project.shape.horizontalflip.md)|
|[ID](Project.shape.id.md)|
|[Left](Project.shape.left.md)|
|[Line](Project.shape.line.md)|
|[LockAspectRatio](Project.shape.lockaspectratio.md)|
|[Name](Project.shape.name.md)|
|[Nodes](Project.shape.nodes.md)|
|[Parent](Project.shape.parent.md)|
|[ParentGroup](Project.shape.parentgroup.md)|
|[Reflection](Project.shape.reflection.md)|
|[Rotation](Project.shape.rotation.md)|
|[Shadow](Project.shape.shadow.md)|
|[ShapeStyle](Project.shape.shapestyle.md)|
|[SoftEdge](Project.shape.softedge.md)|
|[Table](Project.shape.table.md)|
|[TextEffect](Project.shape.texteffect.md)|
|[TextFrame](Project.shape.textframe.md)|
|[TextFrame2](Project.shape.textframe2.md)|
|[ThreeD](Project.shape.threed.md)|
|[Title](Project.shape.title.md)|
|[Top](Project.shape.top.md)|
|[Type](Project.shape.type.md)|
|[VerticalFlip](Project.shape.verticalflip.md)|
|[Vertices](Project.shape.vertices.md)|
|[Visible](Project.shape.visible.md)|
|[Width](Project.shape.width.md)|
|[ZOrderPosition](Project.shape.zorderposition.md)|

## See also


 
[Report Object](Project.report.md)
 
[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]