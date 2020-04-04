---
title: Image.PictureSizeMode Property (Outlook Forms Script)
keywords: olfm10.chm2001930
f1_keywords:
- olfm10.chm2001930
ms.prod: outlook
ms.assetid: f6d435e4-c05c-cc0f-fae2-c5a95743a658
ms.date: 06/08/2017
localization_priority: Normal
---


# Image.PictureSizeMode Property (Outlook Forms Script)

Returns or sets an **Integer** that specifies how to display the background picture on a control. Read/write.


## Syntax

_expression_.**PictureSizeMode**

_expression_ A variable that represents an **Image** object.


## Remarks

The settings for  **PictureSizeMode** are:



|Value|Description|
|:-----|:-----|
|0|Crops any part of the picture that is larger than the form or page (default).|
|1|Stretches the picture to fill the form or page. This setting distorts the picture in either the horizontal or vertical direction.|
|3|Enlarges the picture, but does not distort the picture in either the horizontal or vertical direction.|

The 1 and 3 settings both enlarge the image, but 1 causes distortion. The 1 setting enlarges the image horizontally and vertically until the image reaches the corresponding edges of the container or control. The 3 setting enlarges the image until it reaches either the horizontal or vertical edges of the container or control. If the image reaches the horizontal edges first, any remaining distance to the vertical edges remains blank. If it reaches the vertical edges first, any remaining distance to the horizontal edges remains blank.

Setting the  **PictureSizeMode** property to 2 overrides **[PictureAlignment](Outlook.image.picturealignment.md)**. When  **PictureSizeMode** is set to 2, the picture fills the entire control or container.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]