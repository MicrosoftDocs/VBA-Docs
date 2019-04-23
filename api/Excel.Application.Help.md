---
title: Application.Help method (Excel)
keywords: vbaxl10.chm133146
f1_keywords:
- vbaxl10.chm133146
ms.prod: excel
api_name:
- Excel.Application.Help
ms.assetid: e54291a6-96a5-cc55-72de-f2c1800391e2
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.Help method (Excel)

Displays a Help topic.


## Syntax

_expression_.**Help** (_HelpFile_, _HelpContextID_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _HelpFile_|Optional| **Variant**|The name of the online Help file that you want to display. If this argument isn't specified, Microsoft Excel Help is used.|
| _HelpContextID_|Optional| **Variant**|Specifies the context ID number for the Help topic. If this argument isn't specified, the **Help Topics** dialog box is displayed.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]