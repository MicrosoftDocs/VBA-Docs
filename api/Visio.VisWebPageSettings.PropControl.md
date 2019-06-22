---
title: VisWebPageSettings.PropControl property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.PropControl
ms.assetid: 615e5038-d84d-9527-6987-95f289da77d9
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.PropControl property

Determines whether a shape's custom properties (shape data) are displayed on the webpage. Read/write.


## Syntax

_expression_.**PropControl**

_expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Return value

**Long**


## Remarks

**PropControl** returns non-zero (**True**) if custom properties are set to be displayed on the webpage; otherwise, it returns zero (**False**). The default is **True**.

Set **PropControl** to a non-zero value (**True**) to display custom properties on the webpage; otherwise, set it to zero (**False**). 

If you choose to display custom properties, a **Custom Properties** control appears in the left frame in the browser window, displaying custom properties (shape data) that are associated with a shape when you press Ctrl and choose the shape.

If a shape is part of a group, and both the group and its subshapes have custom properties, the custom properties are displayed in the browser according to the behavior defined in the **Selection** list box on the **Behavior** dialog box (with the group shape selected, choose **Behavior** on the **Format** menu).

The selected behavior determines the display as follows: 

- With **Group only** or **Group first**, Save as Web Page displays the group's custom properties.
    
- With **Members first**, Save as Web Page displays the subshape's custom properties when the mouse pointer moves over a subshape that has custom properties, and group custom properties for those subshapes that do not have custom properties.
    
This behavior can also be set in the SelectMode cell in the Group Properties section of the group shape in the Visio ShapeSheet spreadsheet.

The value of the **PropControl** property corresponds to the setting of the **Details** check box in the **Publishing options** list on the **General** tab of the **Save As Web Page** dialog box (**BackstageButton** tab > **Save As** > **Save as type** list > **Web Page (\*.htm;\*.html)** > **Publish**).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]