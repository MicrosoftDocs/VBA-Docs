---
title: IVBUndoUnit.OnNextAdd method (Visio)
keywords: vis_sdr.chm17360160
f1_keywords:
- vis_sdr.chm17360160
ms.prod: visio
api_name:
- Visio.IVBUndoUnit.OnNextAdd
ms.assetid: a5504398-75a9-06be-346c-3afd85ce708e
ms.date: 06/08/2017
localization_priority: Normal
---


# IVBUndoUnit.OnNextAdd method (Visio)

Notifies an undo unit that another undo unit has been added to the undo stack. Returns **Nothing**.


## Syntax

_expression_.**OnNextAdd**

_expression_ A variable that represents an **[IVBUndoUnit](visio.ivbundounit.md)** object.


## Return value

Nothing


## Remarks

If you are creating an undo unit for your solution, the **OnNextAdd** method is one of the procedures that you must implement. It gets called when the next undo unit in the same undo scope gets added to the undo stack.

When an undo unit receives an **OnNextAdd** notification, it communicates back to the creating object that it can no longer insert data into this undo unit.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]