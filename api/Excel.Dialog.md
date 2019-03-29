---
title: Dialog object (Excel)
keywords: vbaxl10.chm255072
f1_keywords:
- vbaxl10.chm255072
ms.prod: excel
api_name:
- Excel.Dialog
ms.assetid: adabcd3b-fc48-d314-3ae5-f1b2ba148383
ms.date: 03/29/2019
localization_priority: Normal
---


# Dialog object (Excel)

Represents a built-in Microsoft Excel dialog box.


## Remarks

The **Dialog** object is a member of the **[Dialogs](Excel.Dialogs.md)** collection. The **Dialogs** collection contains all the built-in dialog boxes in Microsoft Excel. You cannot create a new built-in dialog box or add one to the collection. The only useful thing that you can do with a **Dialog** object is to use it with the **Show** method to display the corresponding dialog box.

The Microsoft Excel Visual Basic object library includes built-in constants for many of the built-in dialog boxes. Each constant is formed from the prefix "xlDialog" followed by the name of the dialog box. For example, the **Apply Names** dialog box constant is **xlDialogApplyNames**, and the **Find File** dialog box constant is **xlDialogFindFile**. These constants are members of the **[XlBuiltinDialog](Excel.XlBuiltInDialog.md)** enumerated type.


## Example

Use **[Dialogs](Excel.Application.Dialogs.md)** (_index_), where _index_ is a built-in constant identifying the dialog box, to return a single **Dialog** object. The following example runs the built-in **Open** dialog box (**File** menu). The **Show** method returns **True** if Microsoft Excel successfully opens a file; it returns **False** if the user cancels the dialog box.

```vb
dlgAnswer = Application.Dialogs(xlDialogOpen).Show
```


## Methods

- [Show](Excel.Dialog.Show.md)

## Properties

- [Application](Excel.Dialog.Application.md)
- [Creator](Excel.Dialog.Creator.md)
- [Parent](Excel.Dialog.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
