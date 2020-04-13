---
title: Dialogs object (Word)
ms.prod: word
ms.assetid: 8dfa5d8a-bb81-1cdd-853b-3acf9db70aa9
ms.date: 06/08/2017
localization_priority: Normal
---


# Dialogs object (Word)

A collection of  **[Dialog](Word.Dialog.md)** objects in Word. Each **Dialog** object represents a built-in Word dialog box.


## Remarks

Use the **[Dialogs](Word.Application.Dialogs.md)** property to return the **Dialogs** collection. The following example displays the number of available built-in dialog boxes.


```vb
MsgBox Dialogs.Count
```

You cannot create a new built-in dialog box or add one to the **Dialogs** collection. Use **Dialogs** (Index), where Index is the **[WdWordDialog](Word.WdWordDialog.md)** constant that identifies the dialog box, to return a single **Dialog** object. The following example displays the built-in **Open** dialog box.




```vb
dlgAnswer = Dialogs(wdDialogFileOpen).Show
```

For more information, see [Displaying built-in Word dialog boxes](../word/Concepts/Customizing-Word/displaying-built-in-word-dialog-boxes.md).


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]