---
title: FileExportConverter.Description property (Excel)
keywords: vbaxl10.chm863074
f1_keywords:
- vbaxl10.chm863074
ms.prod: excel
api_name:
- Excel.FileExportConverter.Description
ms.assetid: b2bc70da-550b-9286-b534-315ba0916c85
ms.date: 04/26/2019
localization_priority: Normal
---


# FileExportConverter.Description property (Excel)

Returns the description for the file converter. Read-only **String**.


## Syntax

_expression_.**Description**

_expression_ A variable that represents a **[FileExportConverter](Excel.FileExportConverter.md)** object.


## Example

The following example displays the description for the first file converter in the **[FileExportConverters](Excel.FileExportConverters.md)** collection.

```vb
Dim fcTemp As FileExportConverter 
Set fcTemp = FileExportConverter(1) 
 
MsgBox fcTemp.Description
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]