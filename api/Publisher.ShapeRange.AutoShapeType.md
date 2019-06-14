---
title: ShapeRange.AutoShapeType property (Publisher)
keywords: vbapb10.chm2293810
f1_keywords:
- vbapb10.chm2293810
ms.prod: publisher
api_name:
- Publisher.ShapeRange.AutoShapeType
ms.assetid: fa079239-07d8-0783-db34-77ee0f2d5391
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.AutoShapeType property (Publisher)

Returns or sets an **[MsoAutoShapeType](Office.MsoAutoShapeType.md)** constant that specifies a **ShapeRange** object's AutoShape type.


## Syntax

_expression_.**AutoShapeType**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Remarks

The **AutoShapeType** property value can be one of the **MsoAutoShapeType** constants declared in the Microsoft Office type library.

AutoShapes correspond to **Shape** objects, although the **AutoShapeType** property for non-Publisher shapes also return a value. WordArt, OLE, Web Form control, table, and picture frame objects should return **msoShapeMixed** as their **AutoShapeType** property value. Text frames should return **msoShapeRectangle** as their **AutoShapeType** property value.


## Example

This example converts the selected AutoShape object to a lightning bolt if it is a heart and to a 5-point star if it is not. For this example to execute properly, you must have an AutoShape object selected in the active publication.

```vb
Sub ShapeShift() 
 
 Dim srShift As ShapeRange 
 
 Set srShift = Application.ActiveDocument.Selection.ShapeRange 
 If srShift.AutoShapeType = msoShapeHeart Then 
 srShift.AutoShapeType = msoShapeLightningBolt 
 Else 
 srShift.AutoShapeType = msoShape5pointStar 
 End If 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]