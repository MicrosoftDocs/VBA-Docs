---
title: AdvancedPrintOptions.ManualFeedDirection Property (Publisher)
keywords: vbapb10.chm7077929
f1_keywords:
- vbapb10.chm7077929
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.ManualFeedDirection
ms.assetid: 6c241594-d113-c3bd-5669-d3046e824c4e
ms.date: 06/08/2017
localization_priority: Normal
---


# AdvancedPrintOptions.ManualFeedDirection Property (Publisher)

Gets or sets the orientation (landscape or portrait) of how envelopes are fed to the printer's manual feed. Read/write.


## Syntax

 _expression_. **ManualFeedDirection**

 _expression_ A variable that represents an  **AdvancedPrintOptions** object.


## Return value

PbOrientationType


## Remarks

The  **ManualFeedDirection** property setting, in conjunction with the ** [AdvancedPrintOptions.ManualFeedAlign](Publisher.AdvancedPrintOptions.ManualFeedAlign.md)** property setting, corresponds to the **Envelope feed method** setting in the **Envelope Setup** dialog box in the Microsoft Publisher user interface. (On the **File** menu, click **Print Setup**. On the  **Printer Details** tab, click **Advanced Printer Setup**. On the  **Printer Setup Wizard** tab, click **Envelope Setup Dialog**)

Possible values for  **ManualFeedDirection** are **pbOrientationLandscape** (2) and **pbOrientationPortrait** (1).


## See also


 [AdvancedPrintOptions Object](Publisher.AdvancedPrintOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]