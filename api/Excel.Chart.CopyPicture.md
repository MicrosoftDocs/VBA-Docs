---
title: Chart.CopyPicture method (Excel)
keywords: vbaxl10.chm149095
f1_keywords:
- vbaxl10.chm149095
ms.prod: excel
api_name:
- Excel.Chart.CopyPicture
ms.assetid: f69451cd-4be5-982a-58b8-63e0f24e0261
ms.date: 06/08/2017
localization_priority: Priority
---


# Chart.CopyPicture method (Excel)

Copies the selected object to the Clipboard as a picture.


## Syntax

_expression_. `CopyPicture`( `_Appearance_` , `_Format_` , `_Size_` )

_expression_ A variable that represents a [Chart](Excel.Chart-graph-object.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Appearance_|Optional| **[xlPictureAppearance](Excel.XlPictureAppearance.md)**|. Specifies how the picture should be copied. The default value is  **xlScreen**.|
| _Format_|Optional| **[xlCopyPictureFormat](Excel.XlCopyPictureFormat.md)**|. The format of the picture. The default value is  **xlPicture**.|
| _Size_|Optional| **xlPictureAppearance**|The size of the copied picture when the object is a chart on a chart sheet (not embedded on a worksheet). The default value is  **xlPrinter**.|

## See also


[Chart Object](Excel.Chart(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]