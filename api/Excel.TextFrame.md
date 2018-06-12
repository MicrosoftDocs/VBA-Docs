---
title: TextFrame Object (Excel)
keywords: vbaxl10.chm643072
f1_keywords:
- vbaxl10.chm643072
ms.prod: excel
api_name:
- Excel.TextFrame
ms.assetid: 4a6d2201-84b8-d83a-cc13-703da047815e
ms.date: 06/08/2017
---


# TextFrame Object (Excel)

Represents the text frame in a  **[Shape](Excel.Shape.md)** object. Contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame.


## Remarks

Use the  **[TextFrame](Excel.Shape.TextFrame.md)** property to return a **TextFrame** object.


## Example

 The following example adds a rectangle to _myDocument_, adds text to the rectangle, and then sets the margins for the text frame.


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 0, 0, 250, 140).TextFrame 
 .Characters.Text = "Here is some test text" 
 .MarginBottom = 10 
 .MarginLeft = 10 
 .MarginRight = 10 
 .MarginTop = 10 
End With
```


## Methods



|**Name**|
|:-----|
|[Characters](Excel.TextFrame.Characters.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.TextFrame.Application.md)|
|[AutoMargins](Excel.TextFrame.AutoMargins.md)|
|[AutoSize](Excel.TextFrame.AutoSize.md)|
|[Creator](Excel.TextFrame.Creator.md)|
|[HorizontalAlignment](Excel.TextFrame.HorizontalAlignment.md)|
|[HorizontalOverflow](Excel.TextFrame.HorizontalOverflow.md)|
|[MarginBottom](Excel.TextFrame.MarginBottom.md)|
|[MarginLeft](Excel.TextFrame.MarginLeft.md)|
|[MarginRight](Excel.TextFrame.MarginRight.md)|
|[MarginTop](Excel.TextFrame.MarginTop.md)|
|[Orientation](Excel.TextFrame.Orientation.md)|
|[Parent](Excel.TextFrame.Parent.md)|
|[ReadingOrder](Excel.TextFrame.ReadingOrder.md)|
|[VerticalAlignment](Excel.TextFrame.VerticalAlignment.md)|
|[VerticalOverflow](Excel.TextFrame.VerticalOverflow.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
