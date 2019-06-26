---
title: ApplicationSettings.ComplexTextUI property (Visio)
keywords: vis_sdr.chm16251895
f1_keywords:
- vis_sdr.chm16251895
ms.prod: visio
api_name:
- Visio.ApplicationSettings.ComplexTextUI
ms.assetid: b4ea05ad-ef40-6886-de68-c9bfb6826a88
ms.date: 06/08/2017
localization_priority: Normal
---


# ApplicationSettings.ComplexTextUI property (Visio)

Gets whether complex text is displayed in the Microsoft Visio user interface. Read-only.


## Syntax

_expression_.**ComplexTextUI**

_expression_ A variable that represents an **[ApplicationSettings](Visio.ApplicationSettings.md)** object.


## Return value

VisRegionalUIOptions


## Remarks

The following  **VisRegionalUIOptions** constants, which are declared in the Visio type library, show the possible values for the **ComplexTextUI** property.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visRegionalUIOptionsHide**|0|Always hides regional UI.|
| **visRegionalUIOptionsShow**|1|Always shows regional UI.|

The setting of  **ComplexTextUI** is dependent on the settings of three other properties of the **ApplicationSettings** object: **[ApplicationSettings.BIDITextUI](Visio.ApplicationSettings.BIDITextUI.md)**, **[ApplicationSettings.KashidaTextUI](Visio.ApplicationSettings.KashidaTextUI.md)**, and **[ApplicationSettings.SATextUI](Visio.ApplicationSettings.SATextUI.md)**. If any of these properties is set to **visRegionalUIOptionsShow**, **ComplexTextUI** is set to that value as well.

You can determine current language settings by getting the value of the  **[Application.LanguageSettings](Visio.Application.LanguageSettings.md)** property. Or, you can change language settings in the **Microsoft Office Language Preferences** dialog box. (Click **Start**, point to  **All Programs**, point to  **Microsoft Office**, point to  **Microsoft Office Tools**, and then click  **Microsoft Office Language Preferences**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]