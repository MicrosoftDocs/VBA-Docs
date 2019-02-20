---
title: Attachment.PictureTiling property (Access)
keywords: vbaac10.chm13917
f1_keywords:
- vbaac10.chm13917
ms.prod: access
api_name:
- Access.Attachment.PictureTiling
ms.assetid: d7eb8047-ea1d-e864-d2d7-51cd340cbc63
ms.date: 02/07/2019
localization_priority: Normal
---


# Attachment.PictureTiling property (Access)

You can use the **PictureTiling** property to specify whether a background picture is tiled across the entire attachment control. Read/write **Boolean**.


## Syntax

_expression_.**PictureTiling**

_expression_ A variable that represents an **[Attachment](Access.Attachment.md)** object.


## Remarks

The **PictureTiling** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Yes|**True**|The picture is tiled.|
|No|**False**|(Default) The picture isn't tiled.|

You can also set the default for this property by using a control's default control style or the **[DefaultControl](access.form.defaultcontrol.md)** property in Visual Basic.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]