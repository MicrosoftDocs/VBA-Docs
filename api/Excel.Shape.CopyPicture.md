---
title: Shape.CopyPicture method (Excel)
keywords: vbaxl10.chm636127
f1_keywords:
- vbaxl10.chm636127
ms.prod: excel
api_name:
- Excel.Shape.CopyPicture
ms.assetid: 276cd993-18b1-8c5b-3618-95e5b5c9a773
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.CopyPicture method (Excel)

Copies the selected object to the Clipboard as a picture.


## Syntax

_expression_.**CopyPicture** (_Appearance_, _Format_)

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Appearance_|Optional| **Variant**|An **[XlPictureAppearance](Excel.XlPictureAppearance.md)** constant that specifies how the picture should be copied. The default value is **xlScreen**.|
| _Format_|Optional| **Variant**|An **[XlCopyPictureFormat](Excel.XlCopyPictureFormat.md)** constant that specifies the format of the picture. The default value is **xlPicture**.|

## Remarks

If you copy a range, it must be made up of adjacent cells.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]