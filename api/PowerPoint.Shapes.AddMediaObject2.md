---
title: Shapes.AddMediaObject2 method (PowerPoint)
keywords: vbapp10.chm543032
f1_keywords:
- vbapp10.chm543032
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddMediaObject2
ms.assetid: 157499e5-1b90-d85f-b1d8-85a115fc907e
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddMediaObject2 method (PowerPoint)

Replaces deprecated [Shapes.AddMediaObject method (PowerPoint)](PowerPoint.Shapes.AddMediaObject.md). Adds a new media object. 


## Syntax

_expression_. `AddMediaObject2`( `_FileName_`, `_LinkToFile_`, `_SaveWithDocument_`, `_Left_`, `_Top_`, `_Width_`, `_Height_` )

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the file to be added.|
| _LinkToFile_|Optional|**[MSOTRISTATE]**|Indicates whether to link to the file.|
| _SaveWithDocument_|Optional|**[MSOTRISTATE]**|Indicates whether to save the media with the document.|
| _Left_|Optional|**Single**|The distance, in [points](../language/glossary/vbe-glossary.md#point), from the left edge of the slide to the left edge of the media object.|
| _Top_|Optional|**Single**|The distance, in [points](../language/glossary/vbe-glossary.md#point), from the top edge of the slide to the top edge of the media object.|
| _Width_|Optional|**Single**|The width, in [points](../language/glossary/vbe-glossary.md#point), of the media object. Default value is -1.|
| _Height_|Optional|**Single**|The height, in [points](../language/glossary/vbe-glossary.md#point), of the media object. Default value is -1.|

## Return value

 **Shape** object


## Remarks

The default value varies depending whether the new media is an audio or video file, and on file size. If both  _LinkToFile_ and _SaveWithDocument_ are **False**, this method returns an error. At least one must be **True**. Note that the object model allows an object to be both linked and embedded, which is not allowed through the user interface (UI).


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]