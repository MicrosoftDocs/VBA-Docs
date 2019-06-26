---
title: Application.Language property (Visio)
keywords: vis_sdr.chm10013800
f1_keywords:
- vis_sdr.chm10013800
ms.prod: visio
api_name:
- Visio.Application.Language
ms.assetid: 78dc3295-16bd-28fd-43d7-4e6d7924e3be
ms.date: 06/26/2019
localization_priority: Normal
---


# Application.Language property (Visio)

Represents the language ID of the version of the Microsoft Visio instance represented by the parent object. Read/write.


## Syntax

_expression_.**Language**

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Return value

Long


## Remarks

The **Language** property returns the language ID recorded in the object's VERSIONINFO resource. The IDs returned are the standard IDs used by Windows to encode different language versions. For example, the **Language** property returns &H0409 for the U.S. English version of Visio. 

For more information, see [Version information](https://docs.microsoft.com/windows/desktop/menurc/version-information).



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]