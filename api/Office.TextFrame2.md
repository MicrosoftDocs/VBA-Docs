---
title: TextFrame2 Object (Office)
ms.prod: office
api_name:
- Office.TextFrame2
ms.assetid: d2903007-70d4-0b98-e617-96fb2df26975
ms.date: 06/08/2017
---


# TextFrame2 Object (Office)

Represents the text frame in a  **Shape** or **ShapeRange** object. Contains the text in the text frame and exposes properties and methods that control the alignment and anchoring of the text frame.


## Remarks

Use the TextFrame2 property of the Shape and ShapeRange objects to return a TextFrame2 object. 


## Example

The following code adds a rectangle to a slide, adds text to the rectangle, and then sets the margins for the text frame.


```vb
Set pptSlide = ActivePresentation.Slides(1) 
With pptSlide.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame2 
 .TextRange.Text = "Here is some sample text" 
 .MarginBottom = 10 
 .MarginLeft = 10 
 .MarginRight = 10 
 .MarginTop = 10 
End With 

```


## Methods



|**Name**|
|:-----|
|[DeleteText](Office.TextFrame2.DeleteText.md)|

## Properties



|**Name**|
|:-----|
|[Application](Office.TextFrame2.Application.md)|
|[AutoSize](Office.TextFrame2.AutoSize.md)|
|[Column](Office.TextFrame2.Column.md)|
|[Creator](Office.TextFrame2.Creator.md)|
|[HasText](Office.TextFrame2.HasText.md)|
|[HorizontalAnchor](Office.TextFrame2.HorizontalAnchor.md)|
|[MarginBottom](Office.TextFrame2.MarginBottom.md)|
|[MarginLeft](Office.TextFrame2.MarginLeft.md)|
|[MarginRight](Office.TextFrame2.MarginRight.md)|
|[MarginTop](Office.TextFrame2.MarginTop.md)|
|[NoTextRotation](Office.TextFrame2.NoTextRotation.md)|
|[Orientation](Office.TextFrame2.Orientation.md)|
|[Parent](Office.TextFrame2.Parent.md)|
|[PathFormat](Office.TextFrame2.PathFormat.md)|
|[Ruler](Office.TextFrame2.Ruler.md)|
|[TextRange](Office.TextFrame2.TextRange.md)|
|[ThreeD](Office.TextFrame2.ThreeD.md)|
|[VerticalAnchor](Office.TextFrame2.VerticalAnchor.md)|
|[WarpFormat](Office.TextFrame2.WarpFormat.md)|
|[WordArtformat](Office.TextFrame2.WordArtformat.md)|
|[WordWrap](Office.TextFrame2.WordWrap.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
