---
title: Page.ResizeToFitContents method (Visio)
keywords: vis_sdr.chm10950820
f1_keywords:
- vis_sdr.chm10950820
ms.prod: visio
api_name:
- Visio.Page.ResizeToFitContents
ms.assetid: 26b96288-7d8b-a999-ef45-a586110cc8b9
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.ResizeToFitContents method (Visio)

Resizes the page, or the master's page, to fit tightly around the shapes or master that are on it.


## Syntax

_expression_. `ResizeToFitContents`

_expression_ A variable that represents a **[Page](Visio.Page.md)** object.


## Return value

Nothing


## Remarks

After the page is resized, the page height and width and the PinX and PinY values of the shapes or master are typically changed.

Calling the  **ResizeToFitContents** method is the equivalent of selecting **Let Visio expand the page as needed** on the **Page Size** tab in the **Page Setup** dialog box (on the **Design** tab, click **Size**, and then click  **More Page Sizes**).

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this method maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVPage.ResizeToFitContents()**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]