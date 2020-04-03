---
title: Page.AutoSizeDrawing method (Visio)
keywords: vis_sdr.chm10962185
f1_keywords:
- vis_sdr.chm10962185
ms.prod: visio
api_name:
- Visio.AutoSizeDrawing
ms.assetid: 00ae0d14-3268-f6d5-2adb-4653958b6eee
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.AutoSizeDrawing method (Visio)

Automatically resizes the drawing page by adding as many printer-paper-sized tiles as necessary to fit all shapes in the drawing onto the page.


## Syntax

_expression_. `AutoSizeDrawing`

 _expression_ An expression that returns a **[Page](Visio.Page.md)** object.


## Return value

Nothing


## Remarks

If you call the  **AutoSizeDrawing** method when the **Print zoom** setting in the user interface (on the **Print Setup** tab of the **Page Setup** dialog box on the **Design** tab) is set to **Fit to** (a specified number of sheets across and down), Visio returns an error, indicating that it cannot automatically resize that page.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]