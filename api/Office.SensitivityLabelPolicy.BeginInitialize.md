---
title: SensitivityLabelPolicy.BeginInitialize method (Office)
api_name:
- Office.SensitivityLabelPolicy.BeginInitialize
ms.assetid: 45009ace-eaad-4366-9a38-af49a026d05f
ms.date: 03/29/2021
ms.localizationpriority: medium
description: SensitivityLabelPolicy.BeginInitialize method (Office)
---

# SensitivityLabelPolicy.BeginInitialize method (Office)

Begins the **SensitivityLabelPolicy** initialization sequence.

## Syntax

_expression_.**BeginInitialize** ()

_expression_ A variable that represents a **[SensitivityLabelPolicy](Office.SensitivityLabelPolicy.md)** object.


## Return value

String

> [!NOTE]
> A String value that represents the highest supported version of the label policy that should be passed to **CompleteInitialize** method.

## Example

The following example shows the usage of **BeginInitialize** method.

```vb
Dim supportedPolicyVersion As String

supportedPolicyVersion = Application.SensitivityLabelPolicy.BeginInitialize
```


## See also

- [SensitivityLabelPolicy object members](overview/Library-Reference/sensitivitylabelpolicy-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
