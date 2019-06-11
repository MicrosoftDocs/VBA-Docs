---
title: PictureFormat.Height property (Publisher)
keywords: vbapb10.chm3604759
f1_keywords:
- vbapb10.chm3604759
ms.prod: publisher
api_name:
- Publisher.PictureFormat.Height
ms.assetid: d98c76cc-4b75-28b7-5be1-101b372472d5
ms.date: 06/12/2019
localization_priority: Normal
---


# PictureFormat.Height property (Publisher)

Returns a **Variant** that represents the height, in [points](../language/glossary/vbe-glossary.md#point), of the specified picture or OLE object. Read-only.


## Syntax

_expression_.**Height**

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Remarks

The valid range for the **Height** property depends on the size of the application workspace and the position of the object within the workspace. 

For centered objects on non-banner page sizes, the **Height** property may be 0.0 to 50.0 inches. For centered objects on banner page sizes, the **Height** property may be 0.0 to 241.0 inches.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]