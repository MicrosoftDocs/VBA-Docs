---
title: WebNavigationBarSet.HorizontalAlignment property (Publisher)
keywords: vbapb10.chm8519688
f1_keywords:
- vbapb10.chm8519688
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.HorizontalAlignment
ms.assetid: 7d615a5a-793c-fd78-3dca-a268740b67aa
ms.date: 06/18/2019
localization_priority: Normal
---


# WebNavigationBarSet.HorizontalAlignment property (Publisher)

Sets or returns a **[PbWizardNavBarAlignment](publisher.pbwizardnavbaralignment.md)** constant that represents the horizontal alignment of the buttons in a web navigation bar set. Read/write.


## Syntax

_expression_.**HorizontalAlignment**

_expression_ A variable that represents a **[WebNavigationBarSet](Publisher.WebNavigationBarSet.md)** object.


## Return value

PbWizardNavBarAlignment


## Remarks

This property is used to set the way that buttons are displayed in a horizontally oriented web navigation bar set. For example, a **WebNavigationBarSet** object containing 5 links with the **HorizontalButtonCount** property set to 3 and the **HorizontalAlignment** property set to **pbnbAlignRight** aligns the buttons in a grid of 3 columns and 1 row. The first 3 buttons are in the first row and the remaining 2 buttons are in the rightmost columns of the second row.

Returns "Access denied" if **IsHorizontal** = **False** for the specified **WebNavigationBarSet** object. Use the **[ChangeOrientation](Publisher.WebNavigationBarSet.ChangeOrientation.md)** method to set the orientation of the web navigation bar set to horizontal first before setting the **HorizontalAlignment** property.

The **HorizontalAlignment** property value can be set to any of the **PbWizardNavBarAlignment** constants declared in the Microsoft Publisher type library.


## Example

The following example returns the first web navigation bar set from the active document, changes the orientation to horizontal if necessary, sets the **HorizontalButtonCount** property to 3, and then sets the **HorizontalAlignment** property to **pbnbAlignRight**.

```vb
With ActiveDocument.WebNavigationBarSets(1) 
 If .IsHorizontal = False Then 
 .ChangeOrientation pbNavBarOrientHorizontal 
 End If 
 .HorizontalButtonCount = 3 
 .HorizontalAlignment = pbnbAlignRight 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]