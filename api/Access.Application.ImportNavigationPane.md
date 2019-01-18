---
title: Application.ImportNavigationPane method (Access)
keywords: vbaac10.chm12619
f1_keywords:
- vbaac10.chm12619
ms.prod: access
api_name:
- Access.Application.ImportNavigationPane
ms.assetid: 5365ece3-e2da-031c-4e28-89115d48acf8
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ImportNavigationPane method (Access)

Loads a saved Navigation Pane configuration from disk.


## Syntax

_expression_. `ImportNavigationPane`( ` _Path_`, ` _fAppendOnly_` )

_expression_ A variable that represents an [Application](Access.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required|**String**|The path and name of the XML file that contains the Navigation Pane configuration to load. |
| _fAppendOnly_|Optional|**Boolean**|Set to  **True** to append the imported categories to the existing categories. The default value is **False**.|

## See also


[Application Object](Access.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]