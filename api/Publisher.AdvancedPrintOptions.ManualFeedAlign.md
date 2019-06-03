---
title: AdvancedPrintOptions.ManualFeedAlign property (Publisher)
keywords: vbapb10.chm7077928
f1_keywords:
- vbapb10.chm7077928
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.ManualFeedAlign
ms.assetid: 5c2dc0a7-981f-731d-6a85-0971c7e19a62
ms.date: 06/04/2019
localization_priority: Normal
---


# AdvancedPrintOptions.ManualFeedAlign property (Publisher)

Gets or sets the alignment (left, right, or center) of where envelopes are fed to the printer's manual feed. Read/write.


## Syntax

_expression_.**ManualFeedAlign**

_expression_ A variable that represents an **[AdvancedPrintOptions](Publisher.AdvancedPrintOptions.md)** object.


## Return value

**[PbPlacementType](publisher.pbplacementtype.md)**


## Remarks

The **ManualFeedAlign** property setting, in conjunction with the **[ManualFeedDirection](Publisher.AdvancedPrintOptions.ManualFeedDirection.md)** property setting, corresponds to the **Envelope feed method** setting in the **Envelope Setup** dialog box in the Microsoft Publisher user interface. On the **File** menu, choose **Print Setup**. On the **Printer Details** tab, choose **Advanced Printer Setup**. On the **Printer Setup Wizard** tab, choose **Envelope Setup Dialog**.

Possible values for **ManualFeedAlign** are **pbPlacementCenter** (3), **pbPlacementLeft** (1), and **pbPlacementRight** (2).


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]