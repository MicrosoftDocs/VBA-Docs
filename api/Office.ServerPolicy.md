---
title: ServerPolicy object (Office)
keywords: vbaof11.chm278010
f1_keywords:
- vbaof11.chm278010
ms.prod: office
api_name:
- Office.ServerPolicy
ms.assetid: ce2a63d2-5deb-b94b-45d7-ed84e9be7deb
ms.date: 01/23/2019
localization_priority: Normal
---


# ServerPolicy object (Office)

Represents a policy specified for a document type stored on a server running Microsoft Office SharePoint Server.


## Remarks

The **ServerPolicy** object is composed of individual **[PolicyItem](office.policyitem.md)** objects representing the individual policy definitions for the active document.


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

- [ServerPolicy object members](overview/Library-Reference/serverpolicy-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]