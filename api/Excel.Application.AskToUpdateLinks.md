---
title: Application.AskToUpdateLinks property (Excel)
keywords: vbaxl10.chm133079
f1_keywords:
- vbaxl10.chm133079
ms.prod: excel
api_name:
- Excel.Application.AskToUpdateLinks
ms.assetid: 1d04eb45-9dcc-e15f-daf1-a6ce9fa737ae
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.AskToUpdateLinks property (Excel)

**True** if Microsoft Excel asks the user to update links when opening files with links. **False** if links are automatically updated with no dialog box. Read/write **Boolean**.


## Syntax

_expression_.**AskToUpdateLinks**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example sets Microsoft Excel to ask the user to update links whenever a file that contains links is opened.

```vb
Application.AskToUpdateLinks = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]