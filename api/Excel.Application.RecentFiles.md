---
title: Application.RecentFiles property (Excel)
keywords: vbaxl10.chm133170
f1_keywords:
- vbaxl10.chm133170
ms.prod: excel
api_name:
- Excel.Application.RecentFiles
ms.assetid: a64784af-4162-90fc-b955-963a1b1e747f
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.RecentFiles property (Excel)

Returns a **[RecentFiles](Excel.RecentFiles.md)** collection that represents the list of recently used files.


## Syntax

_expression_.**RecentFiles**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example sets the maximum number of files in the list of recently used files to 6.

```vb
Application.RecentFiles.Maximum = 6
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]