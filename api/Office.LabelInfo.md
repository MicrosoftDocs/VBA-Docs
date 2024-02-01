---
title: LabelInfo object (Office)
api_name:
- Office.LabelInfo
ms.assetid: 99e55f38-0884-4458-8c9d-a12fadd7a52d
ms.date: 03/29/2021
ms.localizationpriority: medium
description: LabelInfo object (Office)
---

# LabelInfo object (Office)

Represents the label information data object.

## Remarks

The **LabelInfo** object can be passed to **SetLabel** method of **SensitivityLabel** object.

## Example

The following example shows the usage of members of **LabelInfo** object.

```vb
Sub SetLabelInfo()

 Dim myLabelInfo As Office.LabelInfo
 Set myLabelInfo = Application.ActiveDocument.SensitivityLabel.CreateLabelInfo()
 With myLabelInfo
  .ActionId = "5cc46055-305d-4bc1-8f5f-5edf82231378"
  .AssignmentMethod = MsoAssignmentMethod.PRIVILEGED
  .ContentBits = 4
  .IsEnabled = True
  .Justification = "Some justification needed only if downgrading label."
  .LabelId = "9203368f-916c-4d59-8292-9f1c6a1e8f39"
  .LabelName = "MyLabelName"
  .SetDate = Now()
  .SiteId = "6c15903a-880e-4e17-818a-6cb4f7935615"
 End With

End Sub
```

## See also

- [LabelInfo object members](overview/Library-Reference/labelinfo-members-office.md)
- [SensitivityLabel object members](overview/Library-Reference/sensitivitylabel-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
