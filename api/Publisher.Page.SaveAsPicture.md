---
title: Page.SaveAsPicture method (Publisher)
keywords: vbapb10.chm393272
f1_keywords:
- vbapb10.chm393272
ms.prod: publisher
api_name:
- Publisher.Page.SaveAsPicture
ms.assetid: 9b118126-e072-9516-9863-14ea60264f01
ms.date: 06/11/2019
localization_priority: Normal
---


# Page.SaveAsPicture method (Publisher)

Saves a page as a picture file.


## Syntax

_expression_.**SaveAsPicture** (_FileName_, _pbResolution_)

_expression_ A variable that represents a **[Page](Publisher.Page.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_FileName_|Required| **String**|The path and file name of the new picture file that you want to create. The graphics format that the picture is saved in is determined by the file name extension (such as .jpg or .gif) that you specify.|
|_pbResolution_|Optional| **[PbPictureResolution](Publisher.PbPictureResolution.md)**|The resolution in which you want the picture to be saved. Can be one of the **PbPictureResolution** constants. |

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **SaveAsPicture** method to save the first page of the active publication as a .jpg picture file.

Before running this code, replace `filename.jpg` with a valid file name and the path to a folder on your computer where you have permission to save files.

```vb
Public Sub SaveAsPicture_Example() 
 
 ThisDocument.Pages(1).SaveAsPicture "filename.jpg" 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]