---
title: RecentFiles.Maximum property (Excel)
keywords: vbaxl10.chm172073
f1_keywords:
- vbaxl10.chm172073
ms.prod: excel
api_name:
- Excel.RecentFiles.Maximum
ms.assetid: 24bb3472-8b75-5457-467a-266ed8e5f979
ms.date: 05/11/2019
localization_priority: Normal
---


# RecentFiles.Maximum property (Excel)

Returns or sets the maximum number of files in the list of recently used files. Can be a value from 0 (zero) through 50. Read/write **Long**.


## Syntax

_expression_.**Maximum**

_expression_ A variable that represents a **[RecentFiles](Excel.RecentFiles.md)** object.


## Example

This example sets the maximum number of files in the list of recently used files to 6.

```vb
Application.RecentFiles.Maximum = 6
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]