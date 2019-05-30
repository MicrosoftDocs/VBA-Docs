---
title: Adjustments object (Publisher)
keywords: vbapb10.chm2490367
f1_keywords:
- vbapb10.chm2490367
ms.prod: publisher
api_name:
- Publisher.Adjustments
ms.assetid: a1abecf9-582d-3b5c-8a2c-14c4d260df3a
ms.date: 05/31/2019
localization_priority: Normal
---


# Adjustments object (Publisher)

Contains a collection of adjustment values for the specified AutoShape or WordArt object. 


## Remarks

Each adjustment value represents one way an adjustment handle can be adjusted. Because some adjustment handles can be adjusted in two ways&mdash;for example, some handles can be adjusted both horizontally and vertically&mdash;a shape can have more adjustment values than it has adjustment handles. A shape can have up to eight adjustments.

Use the **[Adjustments](Publisher.Shape.Adjustments.md)** property of the **Shape** object to return an **Adjustments** object. Use **Adjustments** (_index_), where _index_ is the adjustment value's index number, to return a single adjustment value.

Different shapes have different numbers of adjustment values, different kinds of adjustments change the geometry of a shape in different ways, and different kinds of adjustments have different ranges of valid values.

The following table summarizes the ranges of valid adjustment values for different types of adjustments. In most cases, if you specify a value that's beyond the range of valid values, the closest valid value is assigned to the adjustment.

|Type of adjustment|Valid values|
|:-----|:-----|
|Linear (horizontal or vertical)|Generally the value 0.0 represents the left or top edge of the shape and the value 1.0 represents the right or bottom edge of the shape. Valid values correspond to valid adjustments that you can make to the shape manually. For example, if you can only pull an adjustment handle half way across the shape manually, the maximum value for the corresponding adjustment will be 0.5.<br/><br/>For shapes such as callouts, where the values 0.0 and 1.0 represent the limits of the rectangle defined by the starting and ending points of the callout line, negative numbers and numbers greater than 1.0 are valid values.|
|Radial|An adjustment value of 1.0 corresponds to the width of the shape. The maximum value is 0.5, or halfway across the shape.|
|Angle|Values are expressed in degrees. If you specify a value outside the range -180 to 180, it will be normalized to be within that range.|

## Example

The following example adds a right-arrow callout to the active document and sets adjustment values for the callout. Note that although the shape has only three adjustment handles, it has four adjustments. Adjustments three and four both correspond to the handle between the head and neck of the arrow.

```vb
Sub AdjustRightArrowCallout() 
 With ActiveDocument.Pages(1).Shapes.AddShape( _ 
 Type:=msoShapeRightArrowCallout, Left:=72, Top:=72, _ 
 Width:=250, Height:=190).Adjustments 
 .Item(1) = 0.75 'Adjusts width of text box 
 .Item(2) = -0.5 'Adjusts width of arrowhead 
 .Item(3) = 0.8 'Adjusts length of arrowhead 
 .Item(4) = -0.75 'Adjusts width of arrow neck 
 End With 
End Sub
```


## Properties

- [Application](Publisher.Adjustments.Application.md)
- [Count](Publisher.Adjustments.Count.md)
- [Item](Publisher.Adjustments.Item.md)
- [Parent](Publisher.Adjustments.Parent.md)


## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]