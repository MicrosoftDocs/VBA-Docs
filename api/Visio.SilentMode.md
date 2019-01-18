---
title: VisWebPageSettings.SilentMode Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.SilentMode
ms.assetid: 93161e3b-3469-3b86-5143-3ea42229eeea
ms.date: 06/08/2017
localization_priority: Normal
---


# VisWebPageSettings.SilentMode Property (Visio Save As Web)

Determines whether any component of the user interface (either that of Microsoft Visio or that of the browser) is displayed when a drawing is saved as a Web page. Read/write.


## Syntax

 _expression_. **SilentMode**

 _expression_An expression that returns a  ** [VisWebPageSettings](./overview/Visio.md)** object.


## Return value

Long


## Remarks

Set  **SilentMode** to a non-zero value (**True**) to prevent any component of the user interface from appearing when a drawing is saved as a Web page; set it to zero (**False**) to allow dialog boxes to be displayed. The default is  **False**.

Setting the  **SilentMode** property to **True** overrides the setting of the **[OpenBrowser](Visio.OpenBrowser.md)** property and prevents newly created Web pages from opening in the default browser automatically.

To control only the display of dialog boxes in the Visio user interface, use the  **[QuietMode](Visio.QuietMode.md)** property.

If both the  **QuietMode** and **SilentMode** properties are set to **True**, the  **SilentMode** property takes precedence and no user interface components are displayed.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]