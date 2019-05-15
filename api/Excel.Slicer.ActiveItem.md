---
title: Slicer.ActiveItem property (Excel)
keywords: vbaxl10.chm905092
f1_keywords:
- vbaxl10.chm905092
ms.prod: excel
api_name:
- Excel.Slicer.ActiveItem
ms.assetid: ecf95cb2-fb1e-97fc-46a1-2ddcf784a089
ms.date: 05/16/2019
localization_priority: Normal
---


# Slicer.ActiveItem property (Excel)

Returns a **[SlicerItem](Excel.SlicerItem.md)** object that represents the slicer button that is currently in focus for the specified slicer. Read-only.


## Syntax

_expression_.**ActiveItem**

_expression_ A variable that represents a **[Slicer](Excel.Slicer.md)** object.


## Return value

**SlicerItem**


## Remarks

The **ActiveItem** property returns a **SlicerItem** object when the specified slicer has focus and the active control is a button within the slicer (for example, the user can navigate the buttons within the slicer with the keyboard in this state).

The **ActiveItem** property returns **Null** under the following circumstances:

- The specified slicer does not have focus (is not selected or active).
    
- The specified slicer has focus and the whole slicer itself is selected (for example, the user can move the whole slicer around by using the keyboard in this state).
    
- The specified slicer has focus and the active control is the **Clear Filter** button in the header of the slicer.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]