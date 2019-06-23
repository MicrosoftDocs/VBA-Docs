---
title: Application.SetPreviewEnabled method (Visio)
keywords: vis_sdr.chm10062105
f1_keywords:
- vis_sdr.chm10062105
ms.prod: visio
api_name:
- Visio.Application.SetPreviewEnabled
ms.assetid: fa66a148-2eca-85b8-b780-ff077b14d0f2
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SetPreviewEnabled method (Visio)

Turns preview on or off for a gallery in the Microsoft Visio user interface.


## Syntax

_expression_.**SetPreviewEnabled** (_GalleryName_, _OnOrOff_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _GalleryName_|Required| **String**|The name of the gallery for which you want to set the preview status. See Remarks.|
| _OnOrOff_|Required| **Boolean**| **True** to turn preview on for the specified gallery; **False** to turn preview off.|

## Return value

 **Nothing**


## Remarks

For the  _GalleryName_ parameter, you must pass the control ID for the specified gallery. You can find a list of control IDs for all Visio galleries by searching the MSDN library at https://msdn.microsoft.com/library.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]