---
title: FileConverter object (Word)
keywords: vbawd10.chm2457
f1_keywords:
- vbawd10.chm2457
ms.prod: word
api_name:
- Word.FileConverter
ms.assetid: 41af2a9b-75cc-253d-4954-4fb42c88530f
ms.date: 06/08/2017
localization_priority: Normal
---


# FileConverter object (Word)

Represents a file converter that's used to open or save files. The **FileConverter** object is a member of the **FileConverters** collection. The **[FileConverters](Word.fileconverters.md)** collection contains all the installed file converters for opening and saving files.


## Remarks

Use  **FileConverters** (Index), where Index is a class name or index number, to return a single **FileConverter** object. The following example displays the extensions associated with the Microsoft Excel worksheet converter.


```vb
MsgBox FileConverters("MSBiff").Extensions
```

The index number represents the position of the file converter in the **[FileConverters](Word.fileconverters.md)** collection. The following example displays the format name of the first file converter.




```vb
MsgBox FileConverters(1).FormatName
```

You cannot create a new file converter or add one to the **[FileConverters](Word.fileconverters.md)** collection. **FileConverter** objects are added during installation of Microsoft Office or by installing supplemental file converters. Use either the **CanSave** or **CanOpen** property to determine whether a **FileConverter** object can be used to open or save document.

File converters for saving documents are listed in the **Save As** dialog box. File converters for opening documents appear in a dialog box if the **Confirm conversion at Open** check box is selected on the **General** tab in the **Options** dialog box (**Tools** menu).


## Properties



|Name|
|:-----|
|[Application](Word.FileConverter.Application.md)|
|[CanOpen](Word.FileConverter.CanOpen.md)|
|[CanSave](Word.FileConverter.CanSave.md)|
|[ClassName](Word.FileConverter.ClassName.md)|
|[Creator](Word.FileConverter.Creator.md)|
|[Extensions](Word.FileConverter.Extensions.md)|
|[FormatName](Word.FileConverter.FormatName.md)|
|[Name](Word.FileConverter.Name.md)|
|[OpenFormat](Word.FileConverter.OpenFormat.md)|
|[Parent](Word.FileConverter.Parent.md)|
|[Path](Word.FileConverter.Path.md)|
|[SaveFormat](Word.FileConverter.SaveFormat.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]