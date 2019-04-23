---
title: Dialogs object (Excel)
keywords: vbaxl10.chm253072
f1_keywords:
- vbaxl10.chm253072
ms.prod: excel
api_name:
- Excel.Dialogs
ms.assetid: d1d54f0e-6057-92f5-4f4c-254c51e36040
ms.date: 03/29/2019
localization_priority: Normal
---


# Dialogs object (Excel)

A collection of all the **[Dialog](Excel.Dialog.md)** objects in Microsoft Excel.


## Remarks

Each **Dialog** object represents a built-in dialog box. You cannot create a new built-in dialog box or add one to the collection. The only useful thing that you can do with a **Dialog** object is to use it with the **[Show](Excel.Dialog.Show.md)** method to display the corresponding dialog box.

The Microsoft Excel Visual Basic object library includes built-in constants for many of the built-in dialog boxes. Each constant is formed from the prefix "xlDialog" followed by the name of the dialog box. For example, the **Apply Names** dialog box constant is **xlDialogApplyNames**, and the **Find File** dialog box constant is **xlDialogFindFile**. These constants are members of the **[XlBuiltinDialog](Excel.XlBuiltInDialog.md)** enumerated type.


## Example

Use the **[Dialogs](Excel.Application.Dialogs.md)** property of the **Application** object to return the **Dialogs** collection. The following code example displays the number of available built-in Microsoft Excel dialog boxes.

```vb
MsgBox Application.Dialogs.Count
```

<br/>

Use **Dialogs** (_index_), where _index_ is a built-in constant identifying the dialog box, to return a single **Dialog** object. The following example runs the built-in **File Open** dialog box.

```vb
dlgAnswer = Application.Dialogs(xlDialogOpen).Show
```

## Properties

- [Application](Excel.Dialogs.Application.md)
- [Count](Excel.Dialogs.Count.md)
- [Creator](Excel.Dialogs.Creator.md)
- [Item](Excel.Dialogs.Item.md)
- [Parent](Excel.Dialogs.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
