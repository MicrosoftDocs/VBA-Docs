---
title: Shape.AlternativeText property (Visio)
ms.prod: visio
api_name:
- Visio.Shape.AlternativeText
ms.date: 05/08/2019
localization_priority: Normal
---


# Shape.AlternativeText property (Visio)

Returns or sets the alternative text description associated with an object. Read/write.

## Syntax

_expression_.**AlternativeText**

_expression_ A variable that represents a **[Shape](visio.shape.md)** object.

## Return value

String

## Remarks

Use this property to create accessible diagrams by using Visio.

> [!NOTE] 
> Beginning with Microsoft Visio 2016 C2R, you can use **AlternativeText** on **[Page](visio.page.md)**, **Shape**, and **[Master](visio.master.md)** objects. 


## Example

This Visual Basic for Applications (VBA) macro shows how to set and get the **AlternativeText** property of a shape.

```vb
 public Sub ShapeAlternativeText_Example()  
 
     Dim vsoRectangle As Visio.Shape  
     
      'Create a rectangle shape and add alternative text to it. 
    Set vsoRectangle = ActivePage.DrawRectangle(2, 3, 5, 4)   
    vsoRectangle.AlternativeText = "Rectangle Shape alternative text"  
   
     'Get a Cell object from the shape. 
    Set vsoCell = vsoRectangle.Cells("Width")  
  
     'Use the Shape property to get the Shape object. 
    Set vsoShapeFromCell = vsoCell.Shape  
 
     'Use shape's alternative text to verify the proper Shape 
    'object was returned.  
    Debug.Print vsoShapeFromCell.AlternativeText 
 
 End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]