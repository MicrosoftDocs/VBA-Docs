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
| _LinkToFile_|Optional| **VARIANT**|Determines whether the 3D model will be linked to the file from which it was created.|
| _SaveWithDocument_|Optional| **VARIANT**|Determines whether the linked 3D model will be saved with the document into which it is inserted.|
| _Left_|Optional| **VARIANT**|The position (in points) of the upper-left corner of the 3D model relative to the upper-left corner of the document.|
| _Top_|Optional| **VARIANT**|The position (in points) of the upper-left corner of the 3D model relative to the top of the document.|
| _Width_|Optional| **VARIANT**|The width of the 3D model, in points (enter -1 to auto-calculate a width based on the 3D model dimensions).|
| _Height_|Optional| **VARIANT**|The height of the 3D model, in points (enter -1 to auto-calculate a height based on the 3D model dimensions).|

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

This example adds a 3D model created from the file sphere.glb to  _mySheet_. The inserted 3D model is embedded in the active document.


```vb
Set mySheet = Application.ActiveWorkbook.ActiveSheet
Set myShape = mySheet.Shapes.Add3DModel(FileName:="c:\my 3d models\sphere.glb", LinkToFile:=False, SaveWithDocument:=True, Left:=100, Top:=100, Width:=70, Height:=70 )
```


## See also


[Shapes Object](Excel.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]