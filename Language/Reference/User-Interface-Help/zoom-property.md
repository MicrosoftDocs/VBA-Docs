---
title: Zoom property
keywords: fm20.chm2002240
f1_keywords:
- fm20.chm2002240
ms.prod: office
api_name:
- Office.Zoom
ms.assetid: d5230fdb-8332-c136-231b-6e2ab3acaf6a
ms.date: 11/16/2018
localization_priority: Normal
---


# Zoom property

Specifies how much to change the size of a displayed object.

## Syntax

_object_.**Zoom** [= _Integer_ ]

The **Zoom** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Integer_|Optional. The percentage to increase or decrease the displayed image.|

## Remarks

The value of the **Zoom** property specifies a percentage of image enlargement or reduction by which an image display should change. 

Values from 10 to 400 are valid. The value specified is a percentage of the object's original size; thus, a setting of 400 means you want to enlarge the image to four times its original size (or 400 percent), while a setting of 10 means you want to reduce the image to one-tenth of its original size (or 10 percent).

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]