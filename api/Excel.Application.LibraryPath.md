---
title: Application.LibraryPath property (Excel)
keywords: vbaxl10.chm133155
f1_keywords:
- vbaxl10.chm133155
ms.prod: excel
api_name:
- Excel.Application.LibraryPath
ms.assetid: 783efa4a-640b-ab78-2831-da2ecd05558a
ms.date: 06/08/2017
---


# Application.LibraryPath property (Excel)

Returns the path to the Library folder, but without the final separator. Read-only  **String**.


## Syntax

 _expression_. `LibraryPath`

 _expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example opens the file Oscar.xla in the Library folder.


```vb
pathSep = Application.PathSeparator 
f = Application.LibraryPath & pathSep & "Oscar.Xla" 
Workbooks.Open filename:=f
```


## See also


[Application Object](Excel.Application(object).md)

