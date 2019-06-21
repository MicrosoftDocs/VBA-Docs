---
title: InvisibleApp.Version property (Visio)
keywords: vis_sdr.chm17514640
f1_keywords:
- vis_sdr.chm17514640
ms.prod: visio
api_name:
- Visio.InvisibleApp.Version
ms.assetid: fb8929be-b7e7-f8ab-c5a5-5a99dd9b6a89
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.Version property (Visio)

Returns the version of a running Microsoft Visio instance. Read-only.


## Syntax

_expression_.**Version**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

String


## Remarks

Use the  **Version** property of the **InvisibleApp** object to verify the version of a particular Visio instance. This information is helpful if your program requires a particular version. Both the major and minor version numbers are returned. The string returned by Visio is 15.0.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]