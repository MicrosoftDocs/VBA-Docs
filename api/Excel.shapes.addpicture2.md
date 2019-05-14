---
title: Shapes.AddPicture2 method (Excel)
keywords: vbaxl10.chm638097
f1_keywords:
- vbaxl10.chm638097
ms.assetid: 89990ad0-efbc-4262-9ab9-c00c7deac9b5
ms.date: 05/15/2019
ms.prod: excel
localization_priority: Normal
---


# Shapes.AddPicture2 method (Excel)

Creates a picture from an existing file. Returns a **[Shape](Excel.Shape.md)** object that represents the new picture.


## Syntax

_expression_.**AddPicture2** (_FileName_, _LinkToFile_, _SaveWithDocument_, _Left_, _Top_, _Width_, _Height_, _compress_)

_expression_ A variable that represents a **[Shapes](Excel.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The file from which the OLE object is to be created.|
| _LinkToFile_|Required| **[MsoTriState](Office.MsoTriState.md)**|Determines whether the picture will be linked to the file from which it was created.|
| _SaveWithDocument_|Required| **MsoTriState**|Determines whether the linked picture will be saved with the document into which it is inserted. This argument must be **msoTrue** if _LinkToFile_ is **msoFalse**.|
| _Left_|Required| **Single**|The position, measured in [points](../language/glossary/vbe-glossary.md#point), of the left edge of the picture relative to the left edge of the worksheet.|
| _Top_|Required| **Single**|The position, measured in points, of the top edge of the picture relative to the top edge of the worksheet.|
| _Width_|Optional| **Single**|The width of the picture, measured in points.|
| _Height_|Optional| **Single**|The height of the picture, measured in points.|
| _compress_|Optional|**[MsoPictureCompress](overview/Library-Reference/msopicturecompress-enumeration-office.md)** |Determines whether the picture should be compressed when inserted.|

## Return value

**Shape**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
