---
title: Presentation.ApplyTemplate method (PowerPoint)
keywords: vbapp10.chm583007
f1_keywords:
- vbapp10.chm583007
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.ApplyTemplate
ms.assetid: 0340ab20-ae21-996b-63c2-4c0b922dec6e
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.ApplyTemplate method (PowerPoint)

Applies a design template to the specified presentation.


## Syntax

_expression_. `ApplyTemplate`( `_FileName_` )

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|Specifies the name of the design template.|

## Example

This example applies the "Professional" design template to the active presentation.


```vb
Application.ActivePresentation.ApplyTemplate _
    "c:\program files\microsoft office\templates" & _
    "\presentation designs\professional.pot"
```


> [!NOTE] 
> If you refer to an uninstalled presentation design template in a string, a run-time error is generated. The template is not installed automatically regardless of your  **[FeatureInstall](PowerPoint.Application.FeatureInstall.md)** property setting. To use the **ApplyTemplate** method for a template that is not currently installed, you first must install the additional design templates. To do so, install the Additional Design Templates for PowerPoint by running the Microsoft Office installation program (click **Add/Remove Programs** or **Programs and Features** in Windows Control Panel).


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]