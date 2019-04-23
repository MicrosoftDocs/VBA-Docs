---
title: TableBackground object (PowerPoint)
keywords: vbapp10.chm677000
f1_keywords:
- vbapp10.chm677000
ms.prod: powerpoint
api_name:
- PowerPoint.TableBackground
ms.assetid: ba29d6df-f37c-05c1-4e29-8c1766a8aaf4
ms.date: 06/08/2017
localization_priority: Normal
---


# TableBackground object (PowerPoint)

Represents the background associated with a  **Table** object.


## Remarks

Use the  **[Background](PowerPoint.Table.Background.md)** property of a **[Table](PowerPoint.Table.md)** object to return the **TableBackground** object associated with the table.

 To get a **Table** object from an existing shape, use the **Table** property of the **[Shape](PowerPoint.Shape.md)** or **[ShapeRange](PowerPoint.ShapeRange.md)** object that contains the table. You can create a shape that contains a table by using the **[AddTable](PowerPoint.Shapes.AddTable.md)** method of the **[Shapes](PowerPoint.Shapes.md)** collection.

The properties of the  **TableBackground** object return objects that represent various aspects of the formatting associated with a table.


- Use the  **[Fill](PowerPoint.TableBackground.Fill.md)** property to return a **[FillFormat](PowerPoint.FillFormat.md)** object.
    
- Use the  **[Picture](PowerPoint.TableBackground.Picture.md)** property to return a **[PictureFormat](PowerPoint.PictureFormat.md)** object.
    
- Use the  **[Reflection](PowerPoint.TableBackground.Reflection.md)** property to return an **[ReflectionFormat](Office.ReflectionFormat.md)** object.
    
- Use the  **[Shadow](PowerPoint.TableBackground.Shadow.md)** property to return a **[ShadowFormat](PowerPoint.ShadowFormat.md)** object.
    

## Example

The following example shows how to get a  **TableBackground** object and set two of its properties.


```vb
Public Sub TableBackground_Example() 
 
    Dim pptShape As PowerPoint.Shape 
    Dim pptTable As PowerPoint.Table 
    Dim pptTableBackground As PowerPoint.TableBackground 
    Dim pptFillFormat As PowerPoint.FillFormat 
     
    Set pptShape = ActivePresentation.Slides(2).Shapes.AddTable(3, 3) 
    Set pptTable = pptShape.Table 
    Set pptTableBackground = pptTable.Background 
    Set pptFillFormat = pptTableBackground.Fill 
     
    ' Add a patterned fill to the table background 
    pptFillFormat.Patterned (msoPatternSmallGrid) 
     
    ' Add a shadow to the table background 
    pptTableBackground.Shadow.Visible = msoTrue 
     
End Sub
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]