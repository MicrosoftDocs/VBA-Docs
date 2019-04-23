---
title: Application.EnableCheckFileExtensions property (Excel)
keywords: vbaxl10.chm133344
f1_keywords:
- vbaxl10.chm133344
ms.assetid: e518aec5-a261-47ba-a3fd-1da480c82612
ms.date: 04/04/2019
ms.prod: excel
localization_priority: Normal
---


# Application.EnableCheckFileExtensions property (Excel)

**True** to enable the **Tell me if Microsoft Excel isn't the default program for viewing and editing spreadsheets** dialog box. Read/write **Boolean**.


## Syntax

_expression_.**EnableCheckFileExtensions**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example disables the dialog box.


```vb
Application.EnableCheckFileExtensions = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]