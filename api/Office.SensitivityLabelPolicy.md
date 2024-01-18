---
title: SensitivityLabelPolicy object (Office)
api_name:
- Office.SensitivityLabelPolicy
ms.assetid: 6805a678-bf00-4b5c-a5d3-3ca9ee125515
ms.date: 03/29/2021
ms.localizationpriority: medium
description: SensitivityLabelPolicy object (Office)
---


# SensitivityLabelPolicy object (Office)

Represents sensitivity label policy of the user’s organization.

## Remarks

Sensitivity label policy must be initialized before the start of the document otherwise the document will not be able to use the sensitivity label policy set by the organization.

> [!NOTE]
> The organization is identified by using an identity of an Office Account signed into Office.

> [!NOTE]
> Your organization may need a [specific license](/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/microsoft-365-security-compliance-licensing-guidance) to get access to the sensitivity label policy and related functionality.

## Example

The following example shows the usage of members of **SensitivityLabelPolicy** for initializing Office with user's organization policy.


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

Application.SensitivityLabelPolicy.CompleteInitialize myInitInfo

```


## See also

- [SensitivityLabelPolicy object members](overview/Library-Reference/sensitivitylabelpolicy-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
