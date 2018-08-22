---
title: LinkedWindows Collection (VBA Add-In Object Model)
keywords: vbob6.chm1070949
f1_keywords:
- vbob6.chm1070949
ms.prod: office
ms.assetid: 182ee238-c7e5-2d8b-8144-25edd064d1e4
ms.date: 06/08/2017
---


# LinkedWindows Collection (VBA Add-In Object Model)



Contains all linked windows in a [linked window frame](../../Glossary/vbe-glossary.md#linked-window-frame).

## Remarks

Use the  **LinkedWindows** collection to modify the[docked](../../Glossary/vbe-glossary.md#docked) and[linked](../../Glossary/vbe-glossary.md#linked) state of windows in the[development environment](../../Glossary/vbe-glossary.md#development-environment).


 **Important**  Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.


The  **LinkedWindowFrame** property of the **Window** object returns a **Window** object that has a valid **LinkedWindows** collection.
Linked window frames contain all windows that can be linked or docked. This includes all windows except code windows, [designers](../../Glossary/vbe-glossary.md#designer), the [Object Browser](../../Glossary/vbe-glossary.md#Object-Browser) window, and the **Search and Replace** window.
If all the panes from one linked window frame are moved to another window, the linked window frame with no panes is destroyed. However, if all the panes are removed from the main window, it isn't destroyed.
Use the  **Visible** property to check or set the visibility of a window.
You can use the  **Add** method to add a window to the[collection](../../Glossary/vbe-glossary.md#collection) of currently linked windows. A window that is a pane in one linked window frame can be added to another linked window frame. Use the **Remove** method to remove a window from the collection of currently linked windows; this results in the window being unlinked or undocked.
The  **LinkedWindows** collection is used to dock and undock windows from the main window frame.

