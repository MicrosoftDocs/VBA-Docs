---
title: SensitivityLabel.GetLabel method (Office)
api_name:
- Office.SensitivityLabel.GetLabel
ms.assetid: e91a2e48-feff-4468-8c45-d3738b6bd0d8
ms.date: 03/29/2021
ms.localizationpriority: medium
description: SensitivityLabel.GetLabel method (Office)
---

# SensitivityLabel.GetLabel method (Office)

Gets the current label information that exists on the document for the user.

> [!NOTE]
> If the **[SensitivityLabelPolicy.CompleteInitialize](Office.SensitivityLabelPolicy.CompleteInitialize.md)** was called, it gets the label for the user that was passed with **[UserId](Office.SensitivityLabelInitInfo.UserId.md)** otherwise gets the label for the user which is authenticated to the document.

## Syntax

_expression_.**GetLabel** ()

_expression_ A variable that represents a **[SensitivityLabel](Office.SensitivityLabel.md)** object.


## Return value

[LabelInfo](Office.LabelInfo.md)

## Example

The following example shows the usage of **GetLabel** method.

```vb
Dim myLabelInfo As Office.LabelInfo

Set myLabelInfo = ActiveDocument.SensitivityLabel.GetLabel()
```

## See also

- [SensitivityLabel object members](overview/Library-Reference/sensitivitylabel-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
