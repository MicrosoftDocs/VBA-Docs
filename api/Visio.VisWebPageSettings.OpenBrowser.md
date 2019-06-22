---
title: VisWebPageSettings.OpenBrowser property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.OpenBrowser
ms.assetid: 701defdf-9f1c-b136-0af5-48605d255f88
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.OpenBrowser property

Determines whether the webpage opens in the browser after the drawing is exported to a webpage. Read/write.


## Syntax

_expression_.**OpenBrowser**

_expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Return value

**Long**


## Remarks

**OpenBrowser** returns non-zero (**True**) if the webpage opens in the browser after the drawing is exported to a webpage; otherwise, it returns zero (**False**). The default is **True**.

Set **OpenBrowser** to a non-zero value (**True**) to open a webpage in the browser after the drawing is exported to a webpage; otherwise, set it to zero (**False**).

The **OpenBrowser** property corresponds to the **Automatically open Web page in browser** check box on the **General** tab of the **Save As Web Page** dialog box (**BackstageButton** tab > **Save As** > **Save as type** list > **Web Page (\*.htm;\*.html)** > **Publish**).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]