---
title: Chart.CopyPicture method (PowerPoint)
keywords: vbapp10.chm684022
f1_keywords:
- vbapp10.chm684022
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.CopyPicture
ms.assetid: ac8c3f05-3458-8f24-ada8-b89beb52a968
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.CopyPicture method (PowerPoint)

Copies the selected object to the Clipboard as a picture.


## Syntax

_expression_.**CopyPicture** (_Appearance_, _Format_, _Size_)

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Appearance_|Optional|**[XlPictureAppearance](PowerPoint.XlPictureAppearance.md)**|One of the enumeration values that specifies how the picture should be copied. The default is  **xlScreen**.|
| _Format_|Optional|**[XlCopyPictureFormat](PowerPoint.XlCopyPictureFormat.md)**|One of the enumeration values that specifies the format of the picture. The default is  **xlPicture**.|
| _Size_|Optional|**xlPictureAppearance**|One of the enumeration values that specifies the size of the copied picture when the object is a chart on a chart sheet (not embedded on a worksheet). The default is  **xlPrinter**.|

## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]