---
title: ApplicationSettings.KanaFindAndReplace property (Visio)
keywords: vis_sdr.chm16251725
f1_keywords:
- vis_sdr.chm16251725
ms.prod: visio
api_name:
- Visio.ApplicationSettings.KanaFindAndReplace
ms.assetid: 09616d8b-1a81-2c45-c8e5-7b8fa961a4e2
ms.date: 06/08/2017
localization_priority: Normal
---


# ApplicationSettings.KanaFindAndReplace property (Visio)

Gets whether additional options specific to Japanese in the **Find** and **Replace** dialog boxes are available. Read-only.


## Syntax

_expression_.**KanaFindAndReplace**

_expression_ A variable that represents an **[ApplicationSettings](Visio.ApplicationSettings.md)** object.


## Return value

VisRegionalUIOptions


## Remarks

The following **VisRegionalUIOptions** constants, which are declared in the Visio type library, show the possible values for the **KanaFindAndReplace** property.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visRegionalUIOptionsHide**|0|Always hides regional UI.|
| **visRegionalUIOptionsShow**|1|Always shows regional UI.|



> [!NOTE] 
> In Microsoft Office Visio 2003, the **KanaFindAndReplace** property was read/write, and the property setting corresponded to an option on the **Regional** tab of the **Options** dialog box. In Microsoft Office Visio 2007, you can determine current language settings by getting the value of the **[Application.LanguageSettings](Visio.Application.LanguageSettings.md)** property. Or, you can change language settings in the **Microsoft Office Language Settings 2007** dialog box. (Click **Start**, point to **All Programs**, point to **Microsoft Office**, point to **Microsoft Office Tools**, and then click **Microsoft Office 2007 Language Settings**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]