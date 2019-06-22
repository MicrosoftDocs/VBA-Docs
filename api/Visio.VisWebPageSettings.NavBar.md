---
title: VisWebPageSettings.NavBar property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.NavBar
ms.assetid: 5a3245df-d0b6-40c6-5ed9-6d7700e835c8
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.NavBar property

Determines whether the **Go to Page** navigation control is displayed on a webpage. Read/write.


## Syntax

_expression_.**NavBar**

_expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Return value

**Long**


## Remarks

The **NavBar** property returns non-zero (**True**) if the **Go to Page** navigation control is displayed after the drawing is exported to a webpage; otherwise, it returns zero (**False**). The default is **True**.

Set **NavBar** to a non-zero value (**True**) to display the **Go to Page** navigation control after the drawing is exported to a webpage; otherwise, set it to zero (**False**).

This property corresponds to the **Go to Page (navigation control)** check box under **Publishing Options** on the **General** tab of the **Save As Web Page** dialog box (**BackstageButton** tab > **Save As** > **Save as type** list > **Web Page (\*.htm;\*.html)** > **Publish**).

If your Visual Studio solution includes the **Microsoft.Office.Interop.Visio.SaveAsWeb** reference, this property maps to the following types:

- **Microsoft.Office.Interop.Visio.SaveAsWeb.IVisWebPageSettings.NavBar**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]