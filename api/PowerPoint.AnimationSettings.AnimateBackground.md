---
title: AnimationSettings.AnimateBackground property (PowerPoint)
keywords: vbapp10.chm565014
f1_keywords:
- vbapp10.chm565014
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.AnimateBackground
ms.assetid: 929ba50f-23c4-9dea-09fb-fa580715b118
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationSettings.AnimateBackground property (PowerPoint)

If the specified object is an AutoShape, specifies if the shape is animated separately from the text it contains. Read/write.


## Syntax

_expression_. `AnimateBackground`

_expression_ A variable that represents an [AnimationSettings](PowerPoint.AnimationSettings.md) object.


## Remarks

Use the [TextLevelEffect](PowerPoint.AnimationSettings.TextLevelEffect.md)and  **[TextUnitEffect](PowerPoint.AnimationSettings.TextUnitEffect.md)** properties to control the animation of text attached to the specified shape.

If the specified shape is a graph object, the property value is **msoTrue** if the background (the axes and gridlines) of the specified graph object is animated. The property applies only to AutoShapes with text that can be built in more than one step or to graph objects.

If this property is set to  **msoTrue** and the **TextLevelEffect** property is set to **ppAnimateByAllLevels**, the shape and its text are animated simultaneously. If this property is set to **msoTrue** and the **TextLevelEffect** property is set to anything other than **ppAnimateByAllLevels**, the shape is animated immediately before the text is animated.

The effects of setting this property are not apparent unless the specified shape is animated. For a shape to be animated, the  **TextLevelEffect** property for the shape must be set to something other than **ppAnimateLevelNone**, and either the **[Animate](PowerPoint.AnimationSettings.Animate.md)** property must be set to **msoTrue**, or the **[EntryEffect](PowerPoint.AnimationSettings.EntryEffect.md)** property must be set to a constant other than **ppEffectNone**.

The value of the  **AnimateBackground** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The specified shape is not animated separately from the text it contains.|
|**msoTrue**| The specified shape is animated separately from the text it contains.|

## Example

This example creates a rectangle that contains text. The example then specifies that the shape should fly in from the lower-right, that the text should be built from first-level paragraphs, and that the shape should be animated separately from the text it contains. In this example, the  **EntryEffect** property turns on animation.


```vb
Sub AnimateTextBox()

    With ActivePresentation.Slides(1).Shapes.AddShape _
            (Type:=msoShapeRectangle, Left:=50, Top:=200, _
            Width:=200, Height:=200)

        .TextFrame.TextRange = "Reason 1" & Chr(13) & _
        "Reason 2" & Chr(13) & "Reason 3"

        With .AnimationSettings
            .EntryEffect = ppEffectFlyFromBottomRight
            .TextLevelEffect = ppAnimateByFirstLevel
            .TextUnitEffect = ppAnimateByParagraph
            .AnimateBackground = msoTrue
        End With
    End With

End Sub
```


## See also


[AnimationSettings Object](PowerPoint.AnimationSettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]