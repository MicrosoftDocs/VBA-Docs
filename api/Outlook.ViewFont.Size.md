---
title: ViewFont.Size property (Outlook)
keywords: vbaol11.chm2698
f1_keywords:
- vbaol11.chm2698
ms.prod: outlook
api_name:
- Outlook.ViewFont.Size
ms.assetid: 3eecba24-6e4e-637f-bffb-21def66127d8
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewFont.Size property (Outlook)

Returns or sets a  **Long** value that represents the size (in points) of the font in the view. Read-only.


## Syntax

_expression_.**Size**

_expression_ A variable that represents a [ViewFont](Outlook.ViewFont.md) object.


## Remarks

This property can be set to a value between 1 and 127. If this property is set to a value less than 1, the property is set to 1. If this property is set to a value greater than 127, the property is set to 127.

The default value for this property is determined by the operating system.


## See also


[ViewFont Object](Outlook.ViewFont.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]