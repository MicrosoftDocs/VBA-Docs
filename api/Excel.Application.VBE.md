---
title: Application.VBE property (Excel)
keywords: vbaxl10.chm133227
f1_keywords:
- vbaxl10.chm133227
ms.prod: excel
api_name:
- Excel.Application.VBE
ms.assetid: e75dc57a-f9de-beb2-c50c-58445e47d63a
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.VBE property (Excel)

Returns a  **VBE** object that represents the Visual Basic Editor. Read-only.


## Syntax

_expression_. `VBE`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example changes the name of the active Visual Basic project.


```vb
Application.VBE.ActiveVBProject.Name = "TestProject"
```


## See also


[Application Object](Excel.Application(object).md)

