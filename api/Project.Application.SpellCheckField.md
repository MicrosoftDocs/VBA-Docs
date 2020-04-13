---
title: Application.SpellCheckField method (Project)
keywords: vbapj.chm2252
f1_keywords:
- vbapj.chm2252
ms.prod: project-server
api_name:
- Project.Application.SpellCheckField
ms.assetid: 4c5cc4c9-b947-c237-7f7e-0d703bd34352
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SpellCheckField method (Project)

Checks the spelling of text custom fields.


## Syntax

_expression_. `SpellCheckField`( `_FieldName_`, `_EnableSpellCheck_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FieldName_|Optional|**PjSpellingField**|One of the **[PjSpellingField](Project.PjSpellingField.md)** enumeration values.|
| _EnableSpellCheck_|Optional|**Variant**|**True** if spell check is enabled; otherwise, **False**.|

## Return value

 **Boolean**


## Remarks

To check spelling in the entire project, including text custom fields, use the **[SpellingCheck](Project.Application.SpellingCheck.md)** method. The **SpellingCheck** method is equivalent to the **Spelling** command on the **Project** tab of the Ribbon.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]