---
title: Printer.PaperSize property (Publisher)
keywords: vbapb10.chm8978439
f1_keywords:
- vbapb10.chm8978439
ms.prod: publisher
api_name:
- Publisher.Printer.PaperSize
ms.assetid: fa7962fb-3ca0-470a-2337-3193ed0be2aa
ms.date: 06/13/2019
localization_priority: Normal
---


# Printer.PaperSize property (Publisher)

Returns the paper size setting found on the **Publication and Paper Settings** tab in the **Print Setup** dialog box in the Microsoft Publisher user interface (**File** menu). Read-only.


## Syntax

_expression_.**PaperSize**

_expression_ A variable that represents a **[Printer](Publisher.Printer.md)** object.


## Return value

String


## Remarks

If you change the value of either the **PaperHeight** or **PaperWidth** property, the value of the **PaperSize** property changes to Current.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]