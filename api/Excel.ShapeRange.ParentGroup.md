---
title: ShapeRange.ParentGroup property (Excel)
keywords: vbaxl10.chm640131
f1_keywords:
- vbaxl10.chm640131
ms.prod: excel
api_name:
- Excel.ShapeRange.ParentGroup
ms.assetid: b4e8b015-9380-734a-b7e3-74f73c5613fc
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.ParentGroup property (Excel)

Returns a **[Shape](Excel.Shape.md)** object that represents the common parent shape of a child shape or a range of child shapes.


## Syntax

_expression_.**ParentGroup**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Example

In this example, Microsoft Excel adds two shapes to the active worksheet and then removes both shapes by deleting the parent shape of the group.

```vb
Sub ParentGroup() 
 
 Dim pgShape As Shape 
 
 With ActiveSheet.Shapes 
 .AddShape Type:=1, Left:=10, Top:=10, _ 
 Width:=100, Height:=100 
 .AddShape Type:=2, Left:=110, Top:=120, _ 
 Width:=100, Height:=100 
 .Range(Array(1, 2)).Group 
 End With 
 
 ' Using the child shape in the group get the Parent shape. 
 Set pgShape = ActiveSheet.Shapes(1).GroupItems(1).ParentGroup 
 
 MsgBox "The two shapes will now be deleted." 
 
 ' Delete the parent shape. 
 pgShape.Delete 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]