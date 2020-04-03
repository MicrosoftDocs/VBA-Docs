---
title: FormRegion.EnableAutoLayout property (Outlook)
keywords: vbaol11.chm3265
f1_keywords:
- vbaol11.chm3265
ms.prod: outlook
api_name:
- Outlook.FormRegion.EnableAutoLayout
ms.assetid: 24cc737a-0a95-a162-19bb-f2e8e9a73324
ms.date: 06/08/2017
localization_priority: Normal
---


# FormRegion.EnableAutoLayout property (Outlook)

Returns or sets a  **Boolean** that specifies whether to use the automatic layout resizing feature when designing form regions in the forms designer. Read/write


## Syntax

_expression_. `EnableAutoLayout`

_expression_ A variable that represents a [FormRegion](Outlook.FormRegion.md) object.


## Remarks

The automatic layout resizing feature in the forms designer recalculates the form layout dynamically as the form is resized. This feature is only available to form regions in the forms designer.

When this property is  **True**, the forms designer performs automatic layout resizing. When this property is **False**, the forms designer does not perform automatic layout resizing. The default value is **True**.


## See also


[FormRegion Object](Outlook.FormRegion.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]