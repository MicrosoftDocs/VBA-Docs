---
title: TextFrame object (Publisher)
keywords: vbapb10.chm3932159
f1_keywords:
- vbapb10.chm3932159
ms.prod: publisher
api_name:
- Publisher.TextFrame
ms.assetid: 95e88f5a-b3dc-272e-7c1d-5282c97ae11e
ms.date: 06/01/2019
localization_priority: Normal
---


# TextFrame object (Publisher)

Represents the text frame in a **[Shape](Publisher.Shape.md)** object. Contains the text in the text frame and the properties that control the margins and orientation of the text frame.

## Remarks

Use the **[Shape.TextFrame](Publisher.Shape.TextFrame.md)** property to return the **TextFrame** object for a shape. 

The **TextRange** property returns a **[TextRange](Publisher.TextRange.md)** object that represents the range of text inside the specified text frame. 

> [!NOTE] 
> Some shapes do not support attached text (lines, freeforms, pictures, and OLE objects, for example). If you attempt to return or set properties that control text in a text frame for those objects, an error occurs.

Text frames can be linked together so that the text flows from the text frame of one shape into the text frame of another shape. Use the **NextLinkedTextFrame** and **PreviousLinkedTextFrame** properties to link text frames. 

## Example

The following example adds text to the text frame of shape one in the active publication, and then formats the new text.

```vb
Sub AddTextToTextFrame() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 .Text = "My Text" 
 With .Font 
 .Bold = msoTrue 
 .Size = 25 
 .Name = "Arial" 
 End With 
 End With 
End Sub
```

<br/>

Use the **[Shape.HasTextFrame](Publisher.Shape.HasTextFrame.md)** property to determine whether the shape has a text frame, and use the **HasText** property to determine whether the text frame contains text, as shown in the following example.

```vb
Sub GetTextFromTextFrame() 
 Dim shpText As Shape 
 
 For Each shpText In ActiveDocument.Pages(1).Shapes 
 If shpText.HasTextFrame = msoTrue Then 
 With shpText.TextFrame 
 If .HasText Then MsgBox .TextRange.Text 
 End With 
 End If 
 Next 
End Sub
```

<br/>

The following example creates a text box (a rectangle with a text frame) and adds some text to it. It then creates another text box and links the two text frames together so that the text flows from the first text frame into the second one.

```vb
Sub LinkTextBoxes() 
 Dim shpTextBox1 As Shape 
 Dim shpTextBox2 As Shape 
 
 Set shpTextBox1 = ActiveDocument.Pages(1).Shapes.AddTextbox _ 
 (msoTextOrientationHorizontal, 72, 72, 72, 36) 
 shpTextBox1.TextFrame.TextRange.Text = _ 
 "This is some text. This is some more text." 
 
 Set shpTextBox2 = ActiveDocument.Pages(1).Shapes.AddTextbox _ 
 (msoTextOrientationHorizontal, 72, 144, 72, 36) 
 shpTextBox1.TextFrame.NextLinkedTextFrame = shpTextBox2 _ 
 .TextFrame 
End Sub
```


## Methods

- [BreakForwardLink](Publisher.TextFrame.BreakForwardLink.md)
- [ValidLinkTarget](Publisher.TextFrame.ValidLinkTarget.md)

## Properties

- [Application](Publisher.TextFrame.Application.md)
- [AutoFitText](Publisher.TextFrame.AutoFitText.md)
- [Columns](Publisher.TextFrame.Columns.md)
- [ColumnSpacing](Publisher.TextFrame.ColumnSpacing.md)
- [HasNextLink](Publisher.TextFrame.HasNextLink.md)
- [HasPreviousLink](Publisher.TextFrame.HasPreviousLink.md)
- [HasText](Publisher.TextFrame.HasText.md)
- [IncludeContinuedFromPage](Publisher.TextFrame.IncludeContinuedFromPage.md)
- [IncludeContinuedOnPage](Publisher.TextFrame.IncludeContinuedOnPage.md)
- [MarginBottom](Publisher.TextFrame.MarginBottom.md)
- [MarginLeft](Publisher.TextFrame.MarginLeft.md)
- [MarginRight](Publisher.TextFrame.MarginRight.md)
- [MarginTop](Publisher.TextFrame.MarginTop.md)
- [NextLinkedTextFrame](Publisher.TextFrame.NextLinkedTextFrame.md)
- [Orientation](Publisher.TextFrame.Orientation.md)
- [Overflowing](Publisher.TextFrame.Overflowing.md)
- [Parent](Publisher.TextFrame.Parent.md)
- [PreviousLinkedTextFrame](Publisher.TextFrame.PreviousLinkedTextFrame.md)
- [Story](Publisher.TextFrame.Story.md)
- [TextRange](Publisher.TextFrame.TextRange.md)
- [VerticalTextAlignment](Publisher.TextFrame.VerticalTextAlignment.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]