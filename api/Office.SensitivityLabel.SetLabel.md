---
title: SensitivityLabel.SetLabel method (Office)
api_name:
- Office.SensitivityLabel.SetLabel
ms.assetid: 836356a3-e9c4-46c6-a4c3-03f338ac343b
ms.date: 03/29/2021
ms.localizationpriority: medium
description: SensitivityLabel.SetLabel method (Office)
---

# SensitivityLabel.SetLabel method (Office)

Sets the label information on the document for the user.

> [!NOTE]
> If the **[SensitivityLabelPolicy.CompleteInitialize](Office.SensitivityLabelPolicy.CompleteInitialize.md)** was called, it sets the label for the user that was passed with **[UserId](Office.SensitivityLabelInitInfo.UserId.md)** otherwise sets the label for the user which is authenticated to the document.


## Syntax

_expression_.**SetLabel** (_LabelInfo_, _Context_)

_expression_ A variable that represents a **[SensitivityLabel](Office.SensitivityLabel.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _LabelInfo_|Required|**[LabelInfo](Office.LabelInfo.md)**|The label information that needs to be set on the document.|
| _Context_|Required|**Object**|A caller defined context that can be returned with **LabelChanged** event to help ensure that the event was raised because of the **SetLabel** call.|

## See also

- [SensitivityLabel object members](overview/Library-Reference/sensitivitylabel-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
