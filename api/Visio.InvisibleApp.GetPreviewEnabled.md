---
title: InvisibleApp.GetPreviewEnabled Method (Visio)
keywords: vis_sdr.chm17562100
f1_keywords:
- vis_sdr.chm17562100
ms.prod: visio
api_name:
- Visio.InvisibleApp.GetPreviewEnabled
ms.assetid: 4c99a819-9f65-43e6-f162-fe4afc1a3ddf
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.GetPreviewEnabled Method (Visio)

Returns a value that indicates whether preview is enabled for the specified gallery in the Microsoft Visio user interface.


## Syntax

 _expression_. `GetPreviewEnabled`( `_GalleryName_` )

 _expression_ A variable that represents an '[InvisibleApp](Visio.InvisibleApp.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _GalleryName_|Required| **String**|The name of the gallery for which you want to check the preview status. See Remarks.|

## Return value

 **Boolean**


## Remarks

For the  _GalleryName_ parameter, you must pass the control ID for the specified gallery. You can find a list of control IDs for all Visio galleries by searching the MSDN library at https://msdn.microsoft.com/library.


