---
title: Slide.ApplyTemplate method (PowerPoint)
keywords: vbapp10.chm531032
f1_keywords:
- vbapp10.chm531032
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.ApplyTemplate
ms.assetid: ecefec47-697e-57d6-375c-47ccd80268a4
ms.date: 06/08/2017
localization_priority: Normal
---


# Slide.ApplyTemplate method (PowerPoint)

Applies a design template to the specified slide.


## Syntax

_expression_. `ApplyTemplate`( `_FileName_` )

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|Specifies the name of the design template.|

> [!NOTE] 
> If you refer to an uninstalled presentation design template in a string, a run-time error is generated. The template is not installed automatically regardless of your  **[FeatureInstall](PowerPoint.Application.FeatureInstall.md)** property setting. To use the **ApplyTemplate** method for a template that is not currently installed, you first must install the additional design templates. To do so, install the Additional Design Templates for PowerPoint by running the Microsoft Office installation program (click **Add/Remove Programs** or **Programs and Features** in Windows Control Panel).


## See also


[Slide Object](PowerPoint.Slide.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]