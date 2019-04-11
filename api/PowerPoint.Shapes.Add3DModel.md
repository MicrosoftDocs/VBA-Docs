---
title: Shapes.Add3DModel method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.Add3DModel
ms.date: 04/12/2019
localization_priority: Priority
---


# Shapes.Add3DModel method (PowerPoint)

Creates a **[Model3DFormat](PowerPoint.Model3DFormat.md)** object from an existing file. Returns a **[Shape](PowerPoint.Shape.md)** object that represents the new 3D model.


## Syntax

_expression_.**Add3DModel** (_FileName_, _LinkToFile_, _SaveWithDocument_, _Left_, _Top_, _Width_, _Height_)

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The file from which the 3D model object is to be created.|
| _LinkToFile_|Required|**[MsoTriState](Office.MsoTriState.md)**|Determines whether the 3D model will be linked to the file from which it was created.|
| _SaveWithDocument_|Required|**MsoTriState**|Determines whether the linked 3D model will be saved with the document into which it is inserted. This argument must be **msoTrue** if LinkToFile is **msoFalse**.|
| _Left_|Required|**Single**|The position, measured in points, of the left edge of the 3D model relative to the left edge of the slide.|
| _Top_|Required|**Single**|The position, measured in points, of the top edge of the 3D model relative to the top edge of the slide.|
| _Width_|Optional|**Single**|The width of the 3D model, measured in points (enter -1 to auto-calculate a width based on the 3D model dimensions).|
| _Height_|Optional|**Single**|The height of the 3D model, measured in points (enter -1 to auto-calculate a height based on the 3D model dimensions).|

## Return value

Shape


## Example

This example adds a 3D model created from the file Sphere.glb to _mySlide_. The inserted 3D model is embedded in the active document.

```vb
Set mySlide = Application.ActivePresentation.Slides(1) 
Set myShape = mySlide.Shapes.Add3DModel(FileName:="c:\my 3d models\sphere.glb", LinkToFile:=False, SaveWithDocument:=True, Left:=100, Top:=100, Width:=70, Height:=70 )

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]