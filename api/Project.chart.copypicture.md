---
title: Chart.CopyPicture method (Project)
ms.prod: project-server
ms.assetid: 4353ddb2-51f0-a1a4-a472-ec8bbc83b146
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.CopyPicture method (Project)
Copies a selected object to the Clipboard as a picture.

## Syntax

_expression_.**CopyPicture** (_Appearance_, _Format_, _Size_)

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Appearance_|Optional|**Long**|Specifies how the picture should be copied. Can be one of the **Excel.XlPictureAppearance** constants. The default value is **xlScreen** (1).|
| _Format_|Optional|**Long**|Specifies the format of the picture. Can be one of the **Excel.XlCopyPictureFormat** constants. The default value is **xlPicture** (-4147).|
| _Size_|Optional|**Long**|Specifies whether the size of the copied picture should be optimized for a printer or for the screen. Can be one of the **Excel.XlPictureAppearance** constants. The default value is **xlPrinter** (2).|
| _Appearance_|Optional|INT||
| _Format_|Optional|INT||
| _Size_|Optional|INT||

## Return value

 **Nothing**


## See also


[Chart Object](Project.chart.md)
[Copy Method](Project.chart.copy.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]