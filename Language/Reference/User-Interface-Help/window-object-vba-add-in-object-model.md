---
title: Window Object (VBA Add-In Object Model)
keywords: vbob6.chm100108
f1_keywords:
- vbob6.chm100108
ms.prod: office
ms.assetid: 5b9dbbc9-ae3d-b0dc-9fcf-69749749492d
ms.date: 06/08/2017
---


# Window Object (VBA Add-In Object Model)



Represents a window in the [development environment](../../Glossary/vbe-glossary.md#development-environment).

## Remarks

Use the  **Window** object to show, hide, or position windows.


 **Important**  Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.


You can use the  **Close** method to close a window in the **Windows** collection. The **Close** method affects different types of windows as follows:


|**Window**|**Result of using Close method**|
|:-----|:-----|
|Code window|Removes the window from the  **Windows** collection.|
|[Designer](../../Glossary/vbe-glossary.md#Designer)|Removes the window from the  **Windows** collection.|
|**Window** objects of type[linked window frame](../../Glossary/vbe-glossary.md#linked-window-frame)|Windows become unlinked separate windows.|

 **Note**  Using the  **Close** method with code windows and designers actually closes the window. Setting the **Visible** property to **False** hides the window but doesn't close the window. Using the **Close** method with development environment windows, such as the[Project window](../../Glossary/vbe-glossary.md#Project-window) or[Properties window](../../Glossary/vbe-glossary.md#Properties-window), is the same as setting the  **Visible** property to **False**.

You can use the  **SetFocus** method to move the[focus](../../Glossary/vbe-glossary.md#focu) to a window.
You can use the  **Visible** property to return or set the visibility of a window.
To find out what type of window you are working with, you can use the  **Type** property. If you have more than one window of a type, for example, multiple designers, you can use the **Caption** property to determine the window you're working with. You can also find the window you want to work with using the **DesignerWindow** property of the **VBComponent** object or the **Window** property of the **CodePane** object.

