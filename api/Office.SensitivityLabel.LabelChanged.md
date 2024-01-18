---
title: SensitivityLabel.LabelChanged event (Office)
api_name:
- Office.SensitivityLabel.LabelChanged
ms.assetid: 32c64137-10d0-4a27-941a-4cd87664e5be
ms.date: 03/29/2021
ms.localizationpriority: medium
description: SensitivityLabel.LabelChanged event (Office)
---

# SensitivityLabel.LabelChanged event (Office)

Raised when a Label is changed on the document.

## Syntax

_expression_.**LabelChanged** (_OldLabelInfo_, _NewLabelInfo_, _HResult_, _Context_)

_expression_ A variable that represents a **[SensitivityLabel](Office.SensitivityLabel.md)** object.

## Remarks

The **LabelChanged** event is raised after the **SetLabel** is called to indicate the success of **LabelInfo** set operation. If the _HResult_ contains value other than _0_, it indicates **LabelInfo** set operation failure. The _context_ if passed during **SetLabel** is returned here.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_OldLabelInfo_|Required|**LabelInfo**|Previous label information that existed on the document.|
|_NewLabelInfo_|Required|**LabelInfo**|New label information that was applied on the document. |
|_HResult_|Required|**Long**|An integer representing the error code.|
|_Context_|Required|**Object**|The _context_ object that was set with **SetLabel** call. |

## Example

The following example shows the usage of **GetLabel** method.

```vb
Public WithEvents sensitivityLabel As SensitivityLabel

Private Sub sensitivityLabel_LabelChanged(ByVal OldLabelInfo As Office.LabelInfo, ByVal NewLabelInfo As Office.LabelInfo, ByVal HResult As Long, ByVal Context As Object)

 MsgBox "Event raised: " + NewLabelInfo.LabelId

End Sub

Sub SetLabelInfo()

 Set sensitivityLabel = ActiveDocument.SensitivityLabel
 Dim myLabelInfo As Office.LabelInfo
 Set myLabelInfo = sensitivityLabel.CreateLabelInfo()

 With myLabelInfo
  .AssignmentMethod = MsoAssignmentMethod.PRIVILEGED
  .Justification = "Some justification needed only if downgrading label."
  .LabelId = "9203368f-916c-4d59-8292-9f1c6a1e8f39"
  .LabelName = "MyLabelName"
  .SiteId = "6c15903a-880e-4e17-818a-6cb4f7935615"
 End With

 sensitivityLabel.SetLabel myLabelInfo, myLabelInfo

End Sub
```

## See also

- [SensitivityLabel object members](overview/Library-Reference/sensitivitylabel-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
