---
title: FileConverter object (PowerPoint)
keywords: vbapp10.chm680000
f1_keywords:
- vbapp10.chm680000
ms.prod: powerpoint
api_name:
- PowerPoint.FileConverter
ms.assetid: 6baf5bd8-6644-0784-a049-96c3d733043f
ms.date: 06/08/2017
localization_priority: Normal
---


# FileConverter object (PowerPoint)

Represents a file converter that is used to open or save files. The **FileConverter** object is a member of the **FileConverters** collection. The **[FileConverters](PowerPoint.FileConverters.md)** collection contains all the installed file converters for opening and saving files.


## Remarks

Use  **FileConverters** (Index), where Index is a class name or index number, to return a single **FileConverter** object. The following example displays the extensions associated with the Microsoft Excel worksheet converter.


```vb
MsgBox FileConverters("MSBiff").Extensions
```

The index number represents the position of the file converter in the  **[FileConverters](PowerPoint.FileConverters.md)** collection. The following example displays the format name of the first file converter.




```vb
MsgBox FileConverters(1).FormatName
```

You cannot create a new file converter or add one to the  **[FileConverters](PowerPoint.FileConverters.md)** collection. **FileConverter** objects are added during installation of Microsoft Office or by installing supplemental file converters. Use either the **[CanSave](PowerPoint.FileConverter.CanSave.md)** or **[CanOpen](PowerPoint.FileConverter.CanOpen.md)** property to determine whether a **FileConverter** object can be used to open or save document.

File converters for saving documents are listed in the  **Save As** dialog box. File converters for opening documents appear in a dialog box if the **Confirm conversion at Open** check box is selected on the **General** tab in the **Options** dialog box.


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]