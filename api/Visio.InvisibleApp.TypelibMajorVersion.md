---
title: InvisibleApp.TypelibMajorVersion property (Visio)
keywords: vis_sdr.chm17514695
f1_keywords:
- vis_sdr.chm17514695
ms.prod: visio
api_name:
- Visio.InvisibleApp.TypelibMajorVersion
ms.assetid: 22dd9c3f-3c52-29c3-7d99-2230ac3ce90f
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.TypelibMajorVersion property (Visio)

Returns the major version number of the Microsoft Visio type library. Read-only.


## Syntax

_expression_.**TypelibMajorVersion**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

Integer


## Remarks

The major and/or minor version number of the Visio type library will increase whenever the Visio type library is extended. A program can use the  **TypelibMajorVersion** and **TypelibMinorVersion** properties to guarantee that the Visio version it is working with provides support for the features it is using.

Small changes to the Visio type library do not affect the  **Application** object's **Version** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]