---
title: Module.ReplaceLine method (Access)
keywords: vbaac10.chm12279
f1_keywords:
- vbaac10.chm12279
ms.prod: access
api_name:
- Access.Module.ReplaceLine
ms.assetid: 9e267b4a-5c15-a1bc-e2e0-a528871c9268
ms.date: 03/22/2019
localization_priority: Normal
---


# Module.ReplaceLine method (Access)

The **ReplaceLine** method replaces a specified line in a standard module or a class module.


## Syntax

_expression_.**ReplaceLine** (_Line_, _String_)

_expression_ A variable that represents a **[Module](Access.Module.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Line_|Required|**Long**|The number of the line to be replaced.|
| _String_|Required|**String**|The text that is to replace the existing line.|

## Return value

Nothing


## Remarks

Lines in a module are numbered beginning with one. To determine the number of lines in a module, use the **[CountOfLines](Access.Module.CountOfLines.md)** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]