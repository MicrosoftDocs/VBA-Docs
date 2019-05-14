---
title: Shapes.AddPicture method (Excel)
keywords: vbaxl10.chm638082
f1_keywords:
- vbaxl10.chm638082
ms.prod: excel
api_name:
- Excel.Shapes.AddPicture
ms.assetid: 50a46fce-e87d-d5a8-3218-7843788f82bb
ms.date: 05/15/2019
localization_priority: Normal
---


# Shapes.AddPicture method (Excel)

Creates a picture from an existing file. Returns a **[Shape](Excel.Shape.md)** object that represents the new picture.


## Syntax

_expression_.**AddPicture** (_FileName_, _LinkToFile_, _SaveWithDocument_, _Left_, _Top_, _Width_, _Height_)

_expression_ A variable that represents a **[Shapes](Excel.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The file from which the picture is to be created.|
| _LinkToFile_|Required| **[MsoTriState](Office.MsoTriState.md)**| The file to link to. Use **msoFalse** to make the picture an independent copy of the file. Use **msoTrue** to link the picture to the file from which it was created.|
| _SaveWithDocument_|Required| **MsoTriState**|To save the picture with the document. Use **msoFalse** to store only the link information in the document. Use **msoTrue** to save the linked picture with the document into which it's inserted. This argument must be **msoTrue** if _LinkToFile_ is **msoFalse**.|
| _Left_|Required| **Single**|The position (in [points](../language/glossary/vbe-glossary.md#point)) of the upper-left corner of the picture relative to the upper-left corner of the document.|
| _Top_|Required| **Single**|The position (in points) of the upper-left corner of the picture relative to the top of the document.|
| _Width_|Required| **Single**|The width of the picture, in points (enter -1 to retain the width of the existing file).|
| _Height_|Required| **Single**|The height of the picture, in points (enter -1 to retain the height of the existing file).|

## Return value

**Shape**

## Example

This example adds a picture created from the file Music.bmp to _myDocument_. The inserted picture is linked to the file from which it was created and is saved with _myDocument_.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddPicture _ 
    "c:\microsoft office\clipart\music.bmp", _ 
    True, True, 100, 100, 70, 70
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
