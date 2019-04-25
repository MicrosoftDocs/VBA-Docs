---
title: FillFormat.UserTextured method (Excel)
keywords: vbaxl10.chm115010
f1_keywords:
- vbaxl10.chm115010
ms.prod: excel
api_name:
- Excel.FillFormat.UserTextured
ms.assetid: 8c8e7569-50e9-fec5-9c0e-195b26f9394c
ms.date: 04/26/2019
localization_priority: Normal
---


# FillFormat.UserTextured method (Excel)

Fills the specified shape with small tiles of an image. If you want to fill the shape with one large image, use the **[UserPicture](Excel.FillFormat.UserPicture.md)** method.


## Syntax

_expression_.**UserTextured** (_TextureFile_)

_expression_ A variable that represents a **[FillFormat](Excel.FillFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TextureFile_|Required| **String**| The name of the picture file.|

## Example

This example sets the fill format for chart two.

```vb
Charts(2).ChartArea.Fill.UserTextured "brick.gif"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]