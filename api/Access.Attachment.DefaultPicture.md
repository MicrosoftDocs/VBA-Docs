---
title: Attachment.DefaultPicture property (Access)
keywords: vbaac10.chm14001,vbaac10.chm5849
f1_keywords:
- vbaac10.chm14001,vbaac10.chm5849
ms.prod: access
api_name:
- Access.Attachment.DefaultPicture
ms.assetid: 98bc9637-50c9-5831-8170-a32abe5915bc
ms.date: 02/07/2019
localization_priority: Normal
---


# Attachment.DefaultPicture property (Access)

Gets or sets the path and file name of the graphic to be used as a background picture on an attachment control. Read/write **String**.


## Syntax

_expression_.**DefaultPicture**

_expression_ A variable that represents an **[Attachment](Access.Attachment.md)** object.


## Remarks

The **DefaultPicture** property contains _(bitmap)_ or the path and file name of a bitmap or other type of graphic to be displayed.

The default setting is _(none)_. After the graphic is loaded into the object, the property setting is _(bitmap)_ or the path and file name of the graphic. If you delete the path and file name of the graphic from the property setting, the picture is deleted from the object, and the property setting is again _(none)_.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]