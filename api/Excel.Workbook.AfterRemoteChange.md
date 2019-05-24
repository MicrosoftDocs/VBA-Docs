---
title: Workbook.AfterRemoteChange event (Excel)
keywords: vbaxl10.chm504121
f1_keywords:
- vbaxl10.chm504121
ms.prod: excel
api_name:
- Excel.Workbook.AfterRemoteChange
ms.date: 05/25/2019
localization_priority: Normal
---


# Workbook.AfterRemoteChange event (Excel)

Occurs after a remote user's edits to the workbook are merged.

## Syntax

_expression_.**AfterRemoteChange**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Return value

**Nothing**

## Example

This example notifies the user that there was an incoming remote change.

```vb
Private Sub Workbook_AfterRemoteChange()
    'A remote user has made a change to this workbook and that change has been merged.
    'The code in this subroutine will now be run.
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]