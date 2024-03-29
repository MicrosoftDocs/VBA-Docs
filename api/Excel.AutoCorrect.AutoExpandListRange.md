---
title: AutoCorrect.AutoExpandListRange property (Excel)
keywords: vbaxl10.chm545082
f1_keywords:
- vbaxl10.chm545082
api_name:
- Excel.AutoCorrect.AutoExpandListRange
ms.assetid: c91d1e8f-aea2-5db0-a79c-b43eff1e31e4
ms.date: 04/06/2019
ms.localizationpriority: medium
---


# AutoCorrect.AutoExpandListRange property (Excel)

A **Boolean** value indicating whether automatic expansion is enabled for lists. When you type in a cell of an empty row or column next to a list, the list will expand to include that row or column if automatic expansion is enabled. Read/write **Boolean**.


## Syntax

_expression_.**AutoExpandListRange**

_expression_ A variable that represents an **[AutoCorrect](Excel.AutoCorrect(object).md)** object.


## Example

The following example enables automatic expansion of lists when typing in adjacent rows or columns.

```vb
Sub SetAutoExpand 
 
 Application.AutoCorrect.AutoExpandListRange = TRUE 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]