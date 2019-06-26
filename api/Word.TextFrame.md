---
title: TextFrame object (Word)
keywords: vbawd10.chm2482
f1_keywords:
- vbawd10.chm2482
ms.prod: word
api_name:
- Word.TextFrame
ms.assetid: 46f7e410-80d9-9fe9-2224-488b623f8592
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame object (Word)

Represents the text frame in a  **Shape** object. The **TextFrame** object contains the text in the text frame and the properties that control the margins and orientation of the text frame.


## Remarks

Use the  **TextFrame** property to return the **TextFrame** object for a shape. The **TextRange** property returns a **[Range](Word.Range.md)** object that represents the range of text inside the specified text frame. The following example adds text to the text frame of shape one in the active document.


```vb
ActiveDocument.Shapes(1).TextFrame.TextRange.Text = "My Text"
```


> [!NOTE] 
> Some shapes do not support attached text (lines, freeforms, pictures, and OLE objects, for example). If you attempt to return or set properties that control text in a text frame for those objects, an error occurs.

Use the  **HasText** property to determine whether the text frame contains text, as shown in the following example.




```vb
For Each s In ActiveDocument.Shapes 
 With s.TextFrame 
 If .HasText Then MsgBox .TextRange.Text 
 End With 
Next
```

Text frames can be linked together so that the text flows from the text frame of one shape into the text frame of another shape. Use the  **Next** and **Previous** properties to link text frames. The following example creates a text box (a rectangle with a text frame) and adds some text to it. It then creates another text box and links the two text frames together so that the text flows from the first text frame into the second one.




```vb
Set myTB1 = ActiveDocument.Shapes.AddTextbox _ 
 (msoTextOrientationHorizontal, 72, 72, 72, 36) 
myTB1.TextFrame.TextRange = _ 
 "This is some text. This is some more text." 
Set myTB2 = ActiveDocument.Shapes.AddTextbox _ 
 (msoTextOrientationHorizontal, 72, 144, 72, 36) 
myTB1.TextFrame.Next = myTB2.TextFrame
```

Use the  **ContainingRange** property to return a **Range** object that represents the entire story that flows between linked text frames. The following example checks the spelling of the text in TextBox 3 and of any other text that is linked to TextBox 3.




```vb
Set myStory = ActiveDocument.Shapes("TextBox 3") _ 
 .TextFrame.ContainingRange 
myStory.CheckSpelling
```


## Methods



|Name|
|:-----|
|[BreakForwardLink](Word.TextFrame.BreakForwardLink.md)|
|[DeleteText](Word.TextFrame.DeleteText.md)|
|[ValidLinkTarget](Word.TextFrame.ValidLinkTarget.md)|

## Properties



|Name|
|:-----|
|[Application](Word.TextFrame.Application.md)|
|[AutoSize](Word.TextFrame.AutoSize.md)|
|[Column](Word.TextFrame.Column.md)|
|[ContainingRange](Word.TextFrame.ContainingRange.md)|
|[Creator](Word.TextFrame.Creator.md)|
|[HasText](Word.TextFrame.HasText.md)|
|[HorizontalAnchor](Word.TextFrame.HorizontalAnchor.md)|
|[MarginBottom](Word.TextFrame.MarginBottom.md)|
|[MarginLeft](Word.TextFrame.MarginLeft.md)|
|[MarginRight](Word.TextFrame.MarginRight.md)|
|[MarginTop](Word.TextFrame.MarginTop.md)|
|[Next](Word.TextFrame.Next.md)|
|[NoTextRotation](Word.TextFrame.NoTextRotation.md)|
|[Orientation](Word.TextFrame.Orientation.md)|
|[Overflowing](Word.TextFrame.Overflowing.md)|
|[Parent](Word.TextFrame.Parent.md)|
|[PathFormat](Word.TextFrame.PathFormat.md)|
|[Previous](Word.TextFrame.Previous.md)|
|[TextRange](Word.TextFrame.TextRange.md)|
|[ThreeD](Word.TextFrame.ThreeD.md)|
|[VerticalAnchor](Word.TextFrame.VerticalAnchor.md)|
|[WarpFormat](Word.TextFrame.WarpFormat.md)|
|[WordWrap](Word.TextFrame.WordWrap.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
