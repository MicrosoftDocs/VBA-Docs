---
title: VisWebPageSettings.Search property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.Search
ms.assetid: ae7e09e6-7f54-e939-5e5c-12af35c1b303
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.Search property

Determines whether the **Search Pages** control for searching for shapes in a drawing is displayed on a webpage. Read/write.


## Syntax

_expression_.**Search**

_expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Return value

**Long**


## Remarks

Search returns non-zero (**True**) if the **Search Pages** control is displayed after the drawing is exported to a webpage; otherwise, it returns zero (**False**). The default is **True**.

Set **Search** to a non-zero value (**True**) to display the **Search Pages** control after the drawing is exported to a webpage; otherwise, set it to zero (**False**). 

This property corresponds to the **Search Pages** check box under **Publishing Options** on the **General** tab of the **Save As Web Page** dialog box (**BackstageButton** tab > **Save As** > **Save as type** list > **Web Page (\*.htm;\*.html)** > **Publish**).

If your Visual Studio solution includes the **Microsoft.Office.Interop.Visio.SaveAsWeb** reference, this property maps to the following types:

- **Microsoft.Office.Interop.Visio.SaveAsWeb.IVisWebPageSettings.Search**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]