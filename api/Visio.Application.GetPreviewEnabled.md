---
title: Application.GetPreviewEnabled method (Visio)
keywords: vis_sdr.chm10062100
f1_keywords:
- vis_sdr.chm10062100
ms.prod: visio
api_name:
- Visio.Application.GetPreviewEnabled
ms.assetid: 6e0d42b9-f1d4-d8b9-ab9c-7f00ba6c6a9d
ms.date: 06/26/2019
localization_priority: Normal
---


# Application.GetPreviewEnabled method (Visio)

Returns a value that indicates whether preview is enabled for the specified gallery in the Microsoft Visio user interface.


## Syntax

_expression_.**GetPreviewEnabled** (_GalleryName_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _GalleryName_|Required| **String**|The name of the gallery for which you want to check the preview status. See Remarks.|


## Return value

**Boolean**


## Remarks

For the _GalleryName_ parameter, you must pass the control ID for the specified gallery. To download a list of control IDs for all Visio galleries, see [Office 2016 Help Files: Office Fluent User Interface Control Identifiers](https://www.microsoft.com/en-us/download/details.aspx?id=50745).



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]