---
title: MSGWrap.ObjectType Property (Visio)
keywords: vis_sdr.chm16113960
f1_keywords:
- vis_sdr.chm16113960
ms.prod: visio
api_name:
- Visio.MSGWrap.ObjectType
ms.assetid: 30eb3329-a1c4-1c1c-97fc-7f7fdf3ad7e8
ms.date: 06/08/2017
localization_priority: Normal
---


# MSGWrap.ObjectType Property (Visio)

Returns an object's type. Read-only.


## Syntax

 _expression_. `ObjectType`

 _expression_ A variable that represents a [MSGWrap](./Visio.MSGWrap.md) object.


## Return value

Integer


## Remarks

Constants representing object types are prefixed with  **visObjType** and are declared by the Visio type library in **[VisObjectTypes](Visio.VisObjectTypes.md)**.


## Example

This example shows how to use the  **ObjectType** property of a page to iterate recursively through a group and identify the top shape.


```vb
Public Sub ObjectType_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoShapes As Visio.Shapes 
 Dim vsoPage As Visio.Page 
 
 Application.ActiveWindow.Page.Drop Application.Documents.Item("BASIC_U.VSS").Masters.ItemU("Pentagon"), 3#, 8.5 
 
 Application.ActiveWindow.Page.Drop Application.Documents.Item("BASIC_U.VSS").Masters.ItemU("Ellipse"), 3#, 7.625 
 
 Application.ActiveWindow.Page.Drop Application.Documents.Item("BASIC_U.VSS").Masters.ItemU("Rounded rectangle"), 3#, 7# 
 
 Application.ActiveWindow.Page.Drop Application.Documents.Item("BASIC_U.VSS").Masters.ItemU("Circle"), 3#, 6.25 
 
 Application.ActiveWindow.SelectAll 
 
 ActiveWindow.DeselectAll 
 ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemU("Circle"), visSelect 
 ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemU("Rounded rectangle"), visSelect 
 ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemU("Ellipse"), visSelect 
 ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemU("Pentagon"), visSelect 
 ActiveWindow.Selection.Group 
 
 Set vsoPage = ActivePage 
 Set vsoShapes = vsoPage.Shapes 
 Set vsoShape = vsoShapes.Item(2) 
 
 Call GetTopShape(vsoShape) 
 
End Sub 
 
Function GetTopShape(vsoShape As Visio.Shape) As String 
 
 Dim vsoShapeParent As Object 
 Dim vsoShapeParentParent As Object 
 
 Set vsoShapeParent = vsoShape.Parent 
 
 If vsoShapeParent.ObjectType = visObjTypeShape Then 
 
 Set vsoShapeParentParent = vsoShapeParent.Parent 
 
 'If vsoShapeParent's parent isn't a page, keep going up. 
 If vsoShapeParentParent.ObjectType = visObjTypePage Then 
 GetTopShape = vsoShapeParent.Name 
 Else 
 GetTopShape = GetTopShape(vsoShapeParent) 
 End If 
 
 End If 
 
 Debug.Print vsoShapeParent.Name 
 
End Function
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]