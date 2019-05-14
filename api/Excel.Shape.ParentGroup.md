---
title: Shape.ParentGroup property (Excel)
keywords: vbaxl10.chm636139
f1_keywords:
- vbaxl10.chm636139
ms.prod: excel
api_name:
- Excel.Shape.ParentGroup
ms.assetid: 6c43979d-fe16-4093-9eb2-78863230e6d2
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.ParentGroup property (Excel)

Returns a **Shape** object that represents the common parent shape of a child shape or a range of child shapes.


## Syntax

_expression_.**ParentGroup**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


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