---
title: ShapeRange.VerticalFlip property (Publisher)
keywords: vbapb10.chm2293844
f1_keywords:
- vbapb10.chm2293844
ms.prod: publisher
api_name:
- Publisher.ShapeRange.VerticalFlip
ms.assetid: cc3ab3ec-71f6-49fc-0141-505054d6abbb
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.VerticalFlip property (Publisher)

Returns **msoTrue** if the specified shape has been flipped around its vertical axis. Read-only.


## Syntax

_expression_.**VerticalFlip**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Remarks

The **VerticalFlip** property value can be one of the **[MsoTriState](office.msotristate.md)** constants declared in the Microsoft Office type library and shown in the following table.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|The shape has not been flipped around its vertical axis.|
| **msoTriStateMixed**|A return value indicating a combination of **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTriStateToggle**|A set value that switches between **msoTrue** and **msoFalse**.|
| **msoTrue**|The shape has been flipped around its vertical axis.|

## Example

This example restores each shape on the active publication to its original state if it has been flipped horizontally or vertically.

```vb
Sub Flipper() 
 
 Dim shpBall As Shape 
 
 For Each shpBall In ActiveDocument.MasterPages.Item(1).Shapes 
 If shpBall.HorizontalFlip = msoTrue Then shpBall.Flip msoFlipHorizontal 
 If shpBall.VerticalFlip = msoTrue Then shpBall.Flip msoFlipVertical 
 Next 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]