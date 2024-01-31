---
title: SensitivityLabelPolicy.CreateSensitivityLabelInitInfo method (Office)
api_name:
- Office.SensitivityLabelPolicy.CreateSensitivityLabelInitInfo
ms.assetid: 98d78182-a1d1-46f3-9aca-f43606fdc1b0
ms.date: 03/29/2021
ms.localizationpriority: medium
description: SensitivityLabelPolicy.CreateSensitivityLabelInitInfo method (Office)
---

# SensitivityLabelPolicy.CreateSensitivityLabelInitInfo method (Office)

Creates a new **SensitivityLabelInitInfo** object that can be passed to the **CompleteInitialize** method.

## Syntax

_expression_.**CreateSensitivityLabelInitInfo** ()

_expression_ A variable that represents a **[SensitivityLabelPolicy](Office.SensitivityLabelPolicy.md)** object.


## Return value

[SensitivityLabelInitInfo](Office.SensitivityLabelInitInfo.md)

> [!NOTE]
> Returns an empty **SensitivityLabelInitInfo** object. The object should be filled by the user.

## Example

The following example shows the usage of **CreateSensitivityLabelInitInfo** method.

```vb
Dim myInitInfo As Office.SensitivityLabelInitInfo

Set myInitInfo = Application.SensitivityLabelPolicy.CreateSensitivityLabelInitInfo()
```


## See also

- [SensitivityLabelPolicy object members](overview/Library-Reference/sensitivitylabelpolicy-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
