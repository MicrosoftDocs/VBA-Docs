---
title: InvisibleApp.Language property (Visio)
keywords: vis_sdr.chm17513800
f1_keywords:
- vis_sdr.chm17513800
ms.prod: visio
api_name:
- Visio.InvisibleApp.Language
ms.assetid: e8f7408a-5589-d4b4-0e85-95ac714f7e6f
ms.date: 06/26/2019
localization_priority: Normal
---


# InvisibleApp.Language property (Visio)

Represents the language ID of the version of the Microsoft Visio instance represented by the parent object. Read/write.


## Syntax

_expression_.**Language**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

Long


## Remarks

The **Language** property returns the language ID recorded in the object's VERSIONINFO resource. The IDs returned are the standard IDs used by Windows to encode different language versions. For example, the **Language** property returns &H0409 for the U.S. English version of Visio. 

For more information, see [Version information](https://docs.microsoft.com/windows/desktop/menurc/version-information).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]