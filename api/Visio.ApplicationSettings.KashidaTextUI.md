---
title: ApplicationSettings.KashidaTextUI property (Visio)
keywords: vis_sdr.chm16260030
f1_keywords:
- vis_sdr.chm16260030
ms.prod: visio
api_name:
- Visio.ApplicationSettings.KashidaTextUI
ms.assetid: 84270b9c-2ae9-4050-8a68-c04dee0297bf
ms.date: 06/08/2017
localization_priority: Normal
---


# ApplicationSettings.KashidaTextUI property (Visio)

Gets the current setting for display of Kashida text-justification in certain cursive languages. Read-only.


## Syntax

_expression_.**KashidaTextUI**

 _expression_ An expression that returns an **[ApplicationSettings](Visio.ApplicationSettings.md)** object.


## Return value

VisRegionalUIOptions


## Remarks

The following  **VisRegionalUIOptions** constants, which are declared in the Visio type library, show the possible values for the **KashidaTextUI** property.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visRegionalUIOptionsHide**|0|Always hides regional UI.|
| **visRegionalUIOptionsShow**|1|Always shows regional UI|

The setting of the  **KashidaTextUI** property corresponds to the regional options setting in the **Microsoft Office Language Settings 2007** dialog box. (Click **Start**, point to  **All Programs**, point to  **Microsoft Office**, point to  **Microsoft Office Tools**, and then click  **Microsoft Office 2007 Language Settings**. 

The setting of the  **KashidaTextUI** property influences the setting of the **[ApplicationSettings.ComplexTextUI](Visio.ApplicationSettings.ComplexTextUI.md)** property. If **KashidaTextUI** is set to **visRegionalUIOptionsShow**, **ComplexTextUI** is set to that value as well.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]