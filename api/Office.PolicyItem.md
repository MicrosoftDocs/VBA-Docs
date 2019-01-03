---
title: PolicyItem object (Office)
keywords: vbaof11.chm278020
f1_keywords:
- vbaof11.chm278020
ms.prod: office
api_name:
- Office.PolicyItem
ms.assetid: aced7bdc-8ef7-2621-f188-f3c1d44ab6dc
ms.date: 06/08/2017
---


# PolicyItem object (Office)

Represents an item within a  **ServerPolicy** object that contains the settings for one policy.


## Remarks

A policy item cannot exist outside the scope of a policy. Policy items are distinct conditions defined for a document stored on a server running Microsoft Office SharePoint Server.


## Example

The following example lists the name and description of all of the policy items for the active document.


```vb
Sub ListPolicyItems() 
Dim objSrvPolicy As ServerPolicy 
Dim objPolicyItem As PolicyItem 
Dim strPolicyItemList As String 
 
Set objSrvPolicy = ActiveDocument.ServerPolicy 
 
For Each objPolicyItem In objSrvPolicy 
 strPolicyItemList = "Policy Item " &amp; objPolicyItem.Name &amp; " - " &amp; _ 
 objPolicyItem.Description &amp; vbCrLf 
Next 
MsgBox (strPolicyItemList) 
 
End Sub 

```


## Properties



|**Name**|
|:-----|
|[Application](Office.PolicyItem.Application.md)|
|[Creator](Office.PolicyItem.Creator.md)|
|[Data](Office.PolicyItem.Data.md)|
|[Description](Office.PolicyItem.Description.md)|
|[Id](Office.PolicyItem.Id.md)|
|[Name](Office.PolicyItem.Name.md)|
|[Parent](Office.PolicyItem.Parent.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
