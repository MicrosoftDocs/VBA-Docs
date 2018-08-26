---
title: Add Method (VBA Add-In Object Model)
keywords: vbob6.chm1014017
f1_keywords:
- vbob6.chm1014017
ms.prod: office
ms.assetid: 95f4b970-0b0a-a41d-6a7b-8ede6626da67
ms.date: 06/08/2017
---


# Add Method (VBA Add-In Object Model)



<<<<<<< HEAD
Adds an object to a [collection](../../Glossary/vbe-glossary.md).
 **Syntax**
 _object_**.Add(**_component_**)**
=======
Adds an object to a [collection](../../Glossary/vbe-glossary.md#collection).

## Syntax

_object_**.Add(**_component_**)**
>>>>>>> master
The  **Add** syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
<<<<<<< HEAD
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md) that evaluates to an object in the Applies To list.|
| _component_|Required. For the  **LinkedWindows** collection, an object. For the **VBComponents** collection, an enumerated[constant](../../Glossary/vbe-glossary.md) representing a[class module](../../Glossary/vbe-glossary.md), a form, or a [standard module](../../Glossary/vbe-glossary.md). For the  **VBProjects** collection, an enumerated constant representing a project type.|
=======
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the Applies To list.|
| _component_|Required. For the  **LinkedWindows** collection, an object. For the **VBComponents** collection, an enumerated[constant](../../Glossary/vbe-glossary.md#constant) representing a[class module](../../Glossary/vbe-glossary.md#class-module), a form, or a [standard module](../../Glossary/vbe-glossary.md#standard-module). For the  **VBProjects** collection, an enumerated constant representing a project type.|
>>>>>>> master

You can use one of the following constants for the  _component_ argument:


|**Constant**|**Description**|
|:-----|:-----|
|**vbext_ct_ClassModule**|Adds a class module to the collection.|
|**vbext_ct_MSForm**|Adds a form to the collection.|
|**vbext_ct_StdModule**|Adds a standard module to the collection.|
|**vbext_pt_StandAlone**|Adds a standalone project to the collection.|

<<<<<<< HEAD
 **Remarks**
For the  **LinkedWindows** collection, the **Add** method adds a window to the collection of currently[linked windows](../../Glossary/vbe-glossary.md).

 **Note**  You can add a window that is a pane in one [linked window frame](../../Glossary/vbe-glossary.md) to another linked window frame; the window is simply moved from one pane to the other. If the linked window frame that the window was moved from no longer contains any panes, it's destroyed.
=======
## Remarks

For the  **LinkedWindows** collection, the **Add** method adds a window to the collection of currently[linked windows](../../Glossary/vbe-glossary.md#linked-window).

 **Note**  You can add a window that is a pane in one [linked window frame](../../Glossary/vbe-glossary.md#linked-window-frame) to another linked window frame; the window is simply moved from one pane to the other. If the linked window frame that the window was moved from no longer contains any panes, it's destroyed.
>>>>>>> master



 **Important**  Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements generate run-time errors when run on the Macintosh.


<<<<<<< HEAD
For the  **VBComponents** collection, the **Add** method creates a new standard component and adds it to the[project](../../Glossary/vbe-glossary.md).
=======
For the  **VBComponents** collection, the **Add** method creates a new standard component and adds it to the[project](../../Glossary/vbe-glossary.md#project).
>>>>>>> master
For the  **VBComponents** collection, the **Add** method returns a **VBComponent** object. For the **LinkedWindows** collection, the **Add** method returns **Nothing**.
For the  **VBProjects** collection, the **Add** method returns a **VBProject** object and adds a project to the **VBProjects** collection.

