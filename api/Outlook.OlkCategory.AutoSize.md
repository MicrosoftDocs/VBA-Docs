---
title: OlkCategory.AutoSize property (Outlook)
keywords: vbaol11.chm1000439
f1_keywords:
- vbaol11.chm1000439
ms.prod: outlook
api_name:
- Outlook.OlkCategory.AutoSize
ms.assetid: e09b2e18-5fd3-cedc-394c-1080635d1b44
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkCategory.AutoSize property (Outlook)

Returns or sets a **Boolean** that automatically sizes the control to display the entire contents. Read/write.


## Syntax

_expression_.**AutoSize**

_expression_ A variable that represents an [OlkCategory](Outlook.OlkCategory.md) object.


## Remarks

 The default value for this property is **True**.

This control assumes one line unless the contents expand and need more space. If this happens, the control will grow to display the contents and move the remaining form contents down to make space.


## See also


[OlkCategory Object](Outlook.OlkCategory.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]