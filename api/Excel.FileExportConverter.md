---
title: FileExportConverter object (Excel)
keywords: vbaxl10.chm862072
f1_keywords:
- vbaxl10.chm862072
ms.prod: excel
api_name:
- Excel.FileExportConverter
ms.assetid: 299f018e-0dfa-c101-7538-4a285918ac20
ms.date: 03/29/2019
localization_priority: Normal
---


# FileExportConverter object (Excel)

Represents a file converter that is used to save files.

## Remarks

You cannot create a new file converter or add one to the **[FileExportConverters](Excel.FileExportConverters.md)** collection. **FileExportConverter** objects are added during installation of Microsoft Office or by installing supplemental file converters.


## Example

Use **FileExportConverters** (_index_), where _index_ is an integer, to return a single **FileExportConverter** object. The following example displays the extensions associated with the second Microsoft Excel worksheet converter in the collection.

```vb
MsgBox FileExportConverters(2).Extensions

```

<br/>

The index number represents the position of the file converter in the **FileExportConverters** collection. The following example displays the description for the first file converter in the collection.

```vb
MsgBox FileExportConverters(1).Description
```


## Properties

- [Application](Excel.FileExportConverter.Application.md)
- [Creator](Excel.FileExportConverter.Creator.md)
- [Description](Excel.FileExportConverter.Description.md)
- [Extensions](Excel.FileExportConverter.Extensions.md)
- [FileFormat](Excel.FileExportConverter.FileFormat.md)
- [Parent](Excel.FileExportConverter.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]