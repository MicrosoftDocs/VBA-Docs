---
title: SensitivityLabelInitInfo object (Office)
api_name:
- Office.SensitivityLabelInitInfo
ms.assetid: 528946cb-4978-45cc-affe-ebbe080602b0
ms.date: 03/29/2021
ms.localizationpriority: medium
description: SensitivityLabelInitInfo object (Office)
---


# SensitivityLabelInitInfo object (Office)

Represents the sensitivity label policy initialization data object.

## Remarks
The **SensitivityLabelInitInfo** object can be passed to **CompleteInitialization** method of **SensitivityLabelPolicy** object.

## Example

The following example shows the usage of members of **SensitivityLabelInitInfo** object.

```vb
Function GetSensitivityLabelsPolicyXml(policyVersion As String)
 Dim myOrgPolicyInXml as String

 ' Set myOrgPolicyInXml based on your organization’s policy based on policyVersion as XML.
 GetSensitivityLabelsPolicyXml = myOrgPolicyInXml
End Function

Dim supportedPolicyVersion As String
Dim myInitInfo As Office.SensitivityLabelInitInfo

supportedPolicyVersion = Application.SensitivityLabelPolicy.BeginInitialize

Set myInitInfo = Application.SensitivityLabelPolicy.CreateSensitivityLabelInitInfo()
myInitInfo.UserId = “someone@example.com”
myInitInfo.SensitivityLabelsPolicyXml = GetSensitivityLabelsPolicyXml(supportedPolicyVersion)

```

## See also

- [SensitivityLabelInitInfo object members](overview/Library-Reference/sensitivitylabelinitinfo-members-office.md)
- [SensitivityLabelPolicy object members](overview/Library-Reference/sensitivitylabelpolicy-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
