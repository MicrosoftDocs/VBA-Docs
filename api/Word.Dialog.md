---
title: Dialog object (Word)
keywords: vbawd10.chm2488
f1_keywords:
- vbawd10.chm2488
ms.prod: word
api_name:
- Word.Dialog
ms.assetid: f90f6e6d-aaa0-c127-ab37-ca074144eff1
ms.date: 06/08/2017
localization_priority: Normal
---


# Dialog object (Word)

Represents a built-in dialog box. The **Dialog** object is a member of the **[Dialogs](Word.dialogs.md)** collection. The **Dialogs** collection contains all the built-in dialog boxes in Word. You cannot create a new built-in dialog box or add one to the **Dialogs** collection.


## Remarks

Use  **Dialogs** (Index), where Index is a **WdWordDialog** constant that identifies the dialog box, to return a single **Dialog** object. The following example displays and carries out the actions taken in the built-in **Open** dialog box.


```vb
dlgAnswer = Dialogs(wdDialogFileOpen).Show
```

The **WdWordDialog** constants are formed from the prefix "wdDialog" followed by the name of the menu and the dialog box. For example, the constant for the **Page Setup** dialog box is **wdDialogFilePageSetup**, and the constant for the **New** dialog box is **wdDialogFileNew**.

For more information about working with built-in Word dialog boxes, see [Displaying built-in Word dialog boxes](../word/Concepts/Customizing-Word/displaying-built-in-word-dialog-boxes.md).


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]