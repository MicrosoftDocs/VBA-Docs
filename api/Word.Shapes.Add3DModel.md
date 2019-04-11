---
title: Shapes.Add3DModel method (Word)
ms.prod: word
api_name:
- Word.Shapes.Add3DModel
ms.date: 04/12/2019
localization_priority: Priority
---


# Shapes.Add3DModel method (Word)

Adds a 3D model to a drawing canvas. Returns a **[Shape](word.shape.md)** object that represents the 3D model and adds it to the **[CanvasShapes](word.canvasshapes.md)** collection.


## Syntax

_expression_.**Add3DModel** (_FileName_, _LinkToFile_, _SaveWithDocument_, _Left_, _Top_, _Width_, _Height_)

_expression_ Required. A variable that represents a **[Shapes](Word.shapes.md)** collection.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The path and file name of the 3D model.|
| _LinkToFile_|Optional| **Variant**| **True** to link the 3D model to the file from which it was created. **False** to make the 3D model an independent copy of the file. The default value is **False**.|
| _SaveWithDocument_|Optional| **Variant**| **True** to save the linked 3D model with the document. The default value is **False**.|
| _Left_|Optional| **Variant**|The position, measured in points, of the left edge of the new 3D model relative to the drawing canvas.|
| _Top_|Optional| **Variant**|The position, measured in points, of the top edge of the new 3D model relative to the drawing canvas.|
| _Width_|Optional| **Variant**|The width of the 3D model, in points (enter -1 to auto-calculate a width based on the 3D model dimensions).|
| _Height_|Optional| **Variant**|The height of the 3D model, in points (enter -1 to auto-calculate a height based on the 3D model dimensions).|

## Return value

Shape


## Example

This example embeds a 3D model in a newly created drawing canvas in the active document.

```vb
Sub NewCanvasPicture() 
 Dim shpCanvas As Shape 
 
 'Add a drawing canvas to the active document 
 Set shpCanvas = ActiveDocument.Shapes.AddCanvas(Left:=100, Top:=75, Width:=200, Height:=300)
 
 'Add a 3D model to the drawing canvas 
 shpCanvas.CanvasItems.Add3DModel(FileName:="c:\my 3D models\sphere.glb", LinkToFile:=False, SaveWithDocument:=True, Left:=100, Top:=100, Width:=70, Height:=70)
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]