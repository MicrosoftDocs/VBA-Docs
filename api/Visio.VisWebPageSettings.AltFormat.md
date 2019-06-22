---
title: VisWebPageSettings.AltFormat property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.AltFormat
ms.assetid: 60f9af7d-dc5a-d234-976a-51db21473e28
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.AltFormat property

Determines whether a secondary output format for the webpage is defined. Read/write.


## Syntax

_expression_.**AltFormat**

_expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Return value

**Long**


## Remarks

The **AltFormat** property returns non-zero (**True**) if a secondary output format for the webpage is defined; otherwise, it returns zero (**False**). The default is **True**.

Set the **AltFormat** property to a non-zero value (**True**) to enable selection of a secondary output format for the webpage; otherwise, set it to zero (**False**).

The **AltFormat** property is ignored if the primary output format chosen is supported in all browsers by Microsoft Visio 2010. For more information about primary and secondary output formats, see the **[PriFormat](Visio.VisWebPageSettings.PriFormat.md)** and **[SecFormat](Visio.VisWebPageSettings.SecFormat.md)** properties.

The following table shows the compatibility of several browsers with various graphic file types and features.

|Format type|Microsoft Internet Explorer 6 or later|Microsoft Internet Explorer 5 or earlier|Firefox 3 or later|
|:-----|:-----|:-----|:-----|
|XAML|Yes with plug-in|No|Yes with plug-in|
|VML|Yes|Varies|No|
|SVG|Yes with plug-in|Yes with plug-in|Partial|
|PNG|Yes|Yes|Yes|
|GIF|Yes|Yes|Yes|
|JPEG|Yes|Yes|Yes|

The **AltFormat** property corresponds to the **Provide alternate format for older browsers** check box on the **Advanced** tab of the **Save as Web Page** dialog box (**BackstageButton** tab > **Save As** > **Save as type** list > **Web Page (\*.htm;\*.html)** > **Publish** > **Advanced**).


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]