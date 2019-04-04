---
title: Application.PromptForSummaryInfo property (Excel)
keywords: vbaxl10.chm133193
f1_keywords:
- vbaxl10.chm133193
ms.prod: excel
api_name:
- Excel.Application.PromptForSummaryInfo
ms.assetid: 6a7799d7-327f-fdea-9c01-da48cf85655b
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.PromptForSummaryInfo property (Excel)

**True** if Microsoft Excel asks for summary information when files are first saved. Read/write **Boolean**.


## Syntax

_expression_.**PromptForSummaryInfo**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays a prompt that asks for summary information when files are first saved.

```vb
Application.PromptForSummaryInfo = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]