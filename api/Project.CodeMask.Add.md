---
title: CodeMask.Add method (Project)
keywords: vbapj.chm131648
f1_keywords:
- vbapj.chm131648
ms.prod: project-server
api_name:
- Project.CodeMask.Add
ms.assetid: 78a7afaa-1a19-6d64-1341-63955aaff7e3
ms.date: 06/08/2017
localization_priority: Normal
---


# CodeMask.Add method (Project)

Returns a **[CodeMaskLevel](Project.CodeMaskLevel.md)** object.


## Syntax

_expression_.**Add** (_Sequence_, _Length_, _Separator_)

_expression_ A variable that represents a 'CodeMask' object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sequence_|Optional|**Long**|Specifies the type of sequence in the code mask. Can be one of the **[PjCustomOutlineCodeSequence](Project.PjCustomOutlineCodeSequence.md)** constants. The default value is **pjCustomOutlineCodeNumbers**.|
| _Length_|Optional|**Variant**|Specifies the length for a given level in the code mask. Can be the string "Any" or an integer value between 1 and 255. |
| _Separator_|Optional|**String**|The character that separates the level of a code mask from the next code mask. Can be one of the following characters: ".", "-", "+", or "/". |

## Return value

 **CodeMaskLevel**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]