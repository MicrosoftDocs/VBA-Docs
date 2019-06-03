---
title: PolicyItem object (Office)
keywords: vbaof11.chm278020
f1_keywords:
- vbaof11.chm278020
ms.prod: office
api_name:
- Office.PolicyItem
ms.assetid: aced7bdc-8ef7-2621-f188-f3c1d44ab6dc
ms.date: 01/23/2019
localization_priority: Normal
---


# PolicyItem object (Office)

Represents an item within a **[ServerPolicy](office.serverpolicy.md)** object that contains the settings for one policy.


## Remarks

A policy item cannot exist outside the scope of a policy. Policy items are distinct conditions defined for a document stored on a server running Microsoft Office SharePoint Server.


## Example

The following example lists the name and description of all the policy items for the active document.


```vb
Sub ListPolicyItems() 
Dim objSrvPolicy As ServerPolicy 
Dim objPolicyItem As PolicyItem 
Dim strPolicyItemList As String 
 
Set objSrvPolicy = ActiveDocument.ServerPolicy 
 
For Each objPolicyItem In objSrvPolicy 
 strPolicyItemList = "Policy Item " & objPolicyItem.Name & " - " & _ 
 objPolicyItem.Description & vbCrLf 
Next 
MsgBox (strPolicyItemList) 
 
End Sub 

```


## See also

- [PolicyItem object members](overview/Library-Reference/policyitem-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]