---
title: AdvancedPrintOptions.ManualFeedDirection property (Publisher)
keywords: vbapb10.chm7077929
f1_keywords:
- vbapb10.chm7077929
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.ManualFeedDirection
ms.assetid: 6c241594-d113-c3bd-5669-d3046e824c4e
ms.date: 06/04/2019
localization_priority: Normal
---


# AdvancedPrintOptions.ManualFeedDirection property (Publisher)

Gets or sets the orientation (landscape or portrait) of how envelopes are fed to the printer's manual feed. Read/write.


## Syntax

_expression_.**ManualFeedDirection**

_expression_ A variable that represents an **[AdvancedPrintOptions](Publisher.AdvancedPrintOptions.md)** object.


## Return value

**[PbOrientationType](publisher.pborientationtype.md)**


## Remarks

The **ManualFeedDirection** property setting, in conjunction with the **[ManualFeedAlign](Publisher.AdvancedPrintOptions.ManualFeedAlign.md)** property setting, corresponds to the **Envelope feed method** setting in the **Envelope Setup** dialog box in the Microsoft Publisher user interface. On the **File** menu, choose **Print Setup**. On the **Printer Details** tab, choose **Advanced Printer Setup**. On the **Printer Setup Wizard** tab, choose **Envelope Setup Dialog**.

Possible values for **ManualFeedDirection** are **pbOrientationLandscape** (2) and **pbOrientationPortrait** (1).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]