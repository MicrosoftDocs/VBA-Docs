---
title: SensitivityLabel.CreateLabelInfo method (Office)
api_name:
- Office.SensitivityLabel.CreateLabelInfo
ms.assetid: 941c1fb8-1d4b-42e4-a9b8-52cf3309b0ea
ms.date: 03/29/2021
ms.localizationpriority: medium
description: SensitivityLabel.CreateLabelInfo method (Office)
---

# SensitivityLabel.CreateLabelInfo method (Office)

Creates a new **LabelInfo** object that can be passed to **SetLabel** method.

## Syntax

_expression_.**CreateLabelInfo** ()

_expression_ A variable that represents a **[SensitivityLabel](Office.SensitivityLabel.md)** object.


## Return value

[LabelInfo](Office.LabelInfo.md)

> [!NOTE]
> Returns an empty **LabelInfo** object. The object should be filled by the user.

## Example

The following example shows the usage of **CreateLabelInfo** method.

```vb
Dim myLabelInfo As Office.LabelInfo

Set myLabelInfo = ActiveDocument.SensitivityLabel.CreateLabelInfo()
```


## See also

- [SensitivityLabel object members](overview/Library-Reference/sensitivitylabel-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
