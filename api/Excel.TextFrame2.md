---
title: TextFrame2 Object (Excel)
ms.prod: excel
api_name:
- Excel.TextFrame2
ms.assetid: 66ba23e5-9b15-b954-a1db-1bd19b4eb90d
ms.date: 06/08/2017
---


# TextFrame2 Object (Excel)

Represents the text frame in a  **[Shape](Excel.Shape.md)**, **[ShapeRange](Excel.ShapeRange.md)**, or **[ChartFormat](Excel.ChartFormat.md)** object.


## Remarks

This object contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame. Use the  **TextFrame2** property to return a **TextFrame2** object.


## Example

The following example adds a rectangle to  `myDocument`, adds text to the rectangle, and then sets the margins for the text frame.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 0, 0, 250, 140).TextFrame2 
 .TextRange.Text = "Here is some test text" 
 .MarginBottom = 10 
 .MarginLeft = 10 
 .MarginRight = 10 
 .MarginTop = 10 
End With
```


## Methods



|**Name**|
|:-----|
|[DeleteText](Excel.TextFrame2.DeleteText.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.TextFrame2.Application.md)|
|[AutoSize](Excel.TextFrame2.AutoSize.md)|
|[Column](Excel.TextFrame2.Column.md)|
|[Creator](Excel.TextFrame2.Creator.md)|
|[HasText](Excel.TextFrame2.HasText.md)|
|[HorizontalAnchor](Excel.TextFrame2.HorizontalAnchor.md)|
|[MarginBottom](Excel.TextFrame2.MarginBottom.md)|
|[MarginLeft](Excel.TextFrame2.MarginLeft.md)|
|[MarginRight](Excel.TextFrame2.MarginRight.md)|
|[MarginTop](Excel.TextFrame2.MarginTop.md)|
|[NoTextRotation](Excel.TextFrame2.NoTextRotation.md)|
|[Orientation](Excel.TextFrame2.Orientation.md)|
|[Parent](Excel.TextFrame2.Parent.md)|
|[PathFormat](Excel.TextFrame2.PathFormat.md)|
|[Ruler](Excel.TextFrame2.Ruler.md)|
|[TextRange](Excel.TextFrame2.TextRange.md)|
|[ThreeD](Excel.TextFrame2.ThreeD.md)|
|[VerticalAnchor](Excel.TextFrame2.VerticalAnchor.md)|
|[WarpFormat](Excel.TextFrame2.WarpFormat.md)|
|[WordArtformat](Excel.TextFrame2.WordArtformat.md)|
|[WordWrap](Excel.TextFrame2.WordWrap.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
