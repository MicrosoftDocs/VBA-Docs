---
title: Window.Zoom property (Visio)
keywords: vis_sdr.chm11614675
f1_keywords:
- vis_sdr.chm11614675
ms.prod: visio
api_name:
- Visio.Window.Zoom
ms.assetid: 35b6973f-ede6-e731-acf0-59ef03456c47
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.Zoom property (Visio)

Gets or sets the current display size (magnification factor) for a page in a window. Read/write.


## Syntax

_expression_.**Zoom**

_expression_ A variable that represents a **[Window](Visio.Window.md)** object.


## Return value

Double


## Remarks

Valid values range from 0.05 to 9.99 (5% to 999%). The value -1 fits the page into the window. The default value is .67, which is equivalent to the  **Whole Page** setting in the **Zoom** dialog box (on the **View** tab, in the **Zoom** group, click **Zoom**).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]