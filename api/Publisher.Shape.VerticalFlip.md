---
title: Shape.VerticalFlip property (Publisher)
keywords: vbapb10.chm2228308
f1_keywords:
- vbapb10.chm2228308
ms.prod: publisher
api_name:
- Publisher.Shape.VerticalFlip
ms.assetid: b3c7492f-08ee-8fad-102a-8e2a2f69b969
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.VerticalFlip property (Publisher)

Returns **msoTrue** if the specified shape has been flipped around its vertical axis. Read-only.


## Syntax

_expression_.**VerticalFlip**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Remarks

The property value can be one of the **[MsoTriState](office.msotristate.md)** constants declared in the Microsoft Office type library and shown in the following table.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|The shape has not been flipped around its vertical axis.|
| **msoTriStateMixed**|Indicates a combination of **msoTrue** and **msoFalse** for the specified shape range.|
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