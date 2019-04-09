---
title: Shapes.Add3DModel method (Excel)
ms.prod: excel
api_name:
- Excel.Shapes.Add3DModel
ms.date: 04/01/2019
localization_priority: Priority
---


# Shapes.Add3DModel method (Excel)

Creates a 3D model from an existing file. Returns a  **Shape** object that represents the new 3D model.


## Syntax

_expression_.**Add3DModel** ( _Filename_, _LinkToFile_, _SaveWithDocument_, _Left_, _Top_, _Width_, _Height_ )

_expression_ A variable that represents a [Shapes](./Excel.Shapes.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **String**|The file from which the 3D model is to be created.|
| _LinkToFile_|Required| **[MsoTriState](Office.MsoTriState.md)**|Determines whether the 3D model will be linked to the file from which it was created.|
| _SaveWithDocument_|Required| **[MsoTriState](Office.MsoTriState.md)**|Determines whether the linked 3D model will be saved with the document into which it is inserted.|
| _Left_|Required| **Single**|The position (in points) of the upper-left corner of the 3D model relative to the upper-left corner of the document.|
| _Top_|Required| **Single**|The position (in points) of the upper-left corner of the 3D model relative to the top of the document.|
| _Width_|Required| **Single**|The width of the 3D model, in points (enter -1 to auto-calculate a width based on the 3D model dimensions).|
| _Height_|Required| **Single**|The height of the 3D model, in points (enter -1 to auto-calculate a height based on the 3D model dimensions).|

## Return value

Shape


## Remarks


The value of the **LinkToFile** parameter can be one of these **[MsoTriState](Office.MsoTriState.md)** constants.

|Constant|Description|
|:-----|:-----|
| **msoCTrue**|Not supported.|
| **msoFalse**|To make the 3D model an independent copy of the file.|
| **msoTriStateMixed**|Not supported.|
| **msoTriStateToggle**|Not supported.|
| **msoTrue**|To link the 3D model to the file from which it was created.|


The value of the **SaveWithDocument** parameter can be one of these **[MsoTriState](Office.MsoTriState.md)** constants.

|Constant|Description|
|:-----|:-----|
| **msoCTrue**|Not supported.|
| **msoFalse**|To store only the link information in the document.|
| **msoTriStateMixed**|Not supported.|
| **msoTriStateToggle**|Not supported.|
| **msoTrue**|To save the linked 3D model with the document into which it?s inserted. This argument must be **msoTrue** if _LinkToFile_ is **msoFalse**.|


## Example

This example adds a 3D model created from the file sphere.glb to  `myDocument`. The inserted 3D model is linked to the file from which it was created and is saved with  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.Add3DModel _ 
    "c:\my 3d models\sphere.glb", _ 
    True, True, 100, 100, 70, 70
```


## See also


[Shapes Object](Excel.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]