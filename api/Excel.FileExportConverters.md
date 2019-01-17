---
title: FileExportConverters object (Excel)
keywords: vbaxl10.chm864072
f1_keywords:
- vbaxl10.chm864072
ms.prod: excel
api_name:
- Excel.FileExportConverters
ms.assetid: f4b0500e-308a-42e7-a9eb-4a511b8ca754
ms.date: 06/08/2017
localization_priority: Normal
---


# FileExportConverters object (Excel)

A collection of  **[FileExportConverter](Excel.FileExportConverter.md)** objects that represent all the file converters available for saving files.


## Remarks

Use the  **FileExportConverters** property to return the **FileExportConverters** collection.

The  **Add** method is not available for the **FileExportConverters** collection. **FileExportConverter** objects are added during installation of Microsoft Office or by installing supplemental converters.


## Example

Use  **FileExportConverters** ( _Index_ ), where _Index_ is an integer, to return a single **FileExportConverter** object. The following example displays the extensions associated with the second Microsoft Excel worksheet converter in the collection.


```vb
MsgBox FileExportConverters(2).Extensions
```

The index number represents the position of the file converter in the  **FileExportConverters** collection. The following example displays the description for the first file converter in the collection.




```vb
MsgBox FileExportConvters(1).Description
```


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)


