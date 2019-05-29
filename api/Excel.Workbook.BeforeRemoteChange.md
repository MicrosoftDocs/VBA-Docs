---
title: Workbook.BeforeRemoteChange event (Excel)
keywords: vbaxl10.chm504120
f1_keywords:
- vbaxl10.chm504120
ms.prod: excel
api_name:
- Excel.Workbook.BeforeRemoteChange
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.BeforeRemoteChange event (Excel)

Occurs before a remote user's edits to the workbook are merged.

## Syntax

_expression_.**BeforeRemoteChange**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Return value

**Nothing**

## Example

This example notifies the user that there is an incoming remote change.

```vb
Private Sub Workbook_BeforeRemoteChange()
    'A remote user has made a change to this workbook.
    'The code in this subroutine will be run before those changes are merged.
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]