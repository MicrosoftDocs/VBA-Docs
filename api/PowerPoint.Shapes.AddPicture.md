---
title: Shapes.AddPicture method (PowerPoint)
keywords: vbapp10.chm543010
f1_keywords:
- vbapp10.chm543010
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddPicture
ms.assetid: af432432-b09b-3ca6-d392-132bd78251c7
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddPicture method (PowerPoint)

Creates a picture from an existing file. Returns a **[Shape](PowerPoint.Shape.md)** object that represents the new picture.


## Syntax

_expression_. `AddPicture`( `_FileName_`, `_LinkToFile_`, `_SaveWithDocument_`, `_Left_`, `_Top_`, `_Width_`, `_Height_` )

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The file from which the OLE object is to be created.|
| _LinkToFile_|Required|**[MsoTriState](Office.MsoTriState.md)**|Determines whether the picture will be linked to the file from which it was created.|
| _SaveWithDocument_|Required|**[MsoTriState](Office.MsoTriState.md)**|Determines whether the linked picture will be saved with the document into which it is inserted. This argument must be  **msoTrue** if LinkToFile is **msoFalse**.|
| _Left_|Required|**Single**|The position, measured in points, of the left edge of the picture relative to the left edge of the slide.|
| _Top_|Required|**Single**|The position, measured in points, of the top edge of the picture relative to the top edge of the slide.|
| _Width_|Optional|**Single**|The width of the picture, measured in points.|
| _Height_|Optional|**Single**|The height of the picture, measured in points.|

## Return value

Shape


## Example

This example adds a picture created from the file Music.bmp to myDocument. The inserted picture is linked to the file from which it was created and is saved with myDocument.


```vb
Set myDocument = ActivePresentation.Slides(1) 
myDocument.Shapes.AddPicture FileName:="c:\microsoft office\" & _ 
    "clipart\music.bmp", LinkToFile:=msoTrue, SaveWithDocument:=msoTrue, _ 
    Left:=100, Top:=100, Width:=70, Height:=70
```


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
