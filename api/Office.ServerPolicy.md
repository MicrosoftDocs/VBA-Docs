---
title: ServerPolicy object (Office)
keywords: vbaof11.chm278010
f1_keywords:
- vbaof11.chm278010
ms.prod: office
api_name:
- Office.ServerPolicy
ms.assetid: ce2a63d2-5deb-b94b-45d7-ed84e9be7deb
ms.date: 06/08/2017
---


# ServerPolicy object (Office)

Represents a policy specified for a document type stored on a server running Microsoft Office SharePoint Server.


## Remarks

The  **ServerPolicy** object is composed of individual **PolicyItem** objects representing the individual policy definitions for the active document.


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



|Name|
|:-----|
|[Application](Office.ServerPolicy.Application.md)|
|[BlockPreview](Office.ServerPolicy.BlockPreview.md)|
|[Count](Office.ServerPolicy.Count.md)|
|[Creator](Office.ServerPolicy.Creator.md)|
|[Description](Office.ServerPolicy.Description.md)|
|[Id](Office.ServerPolicy.Id.md)|
|[Item](Office.ServerPolicy.Item.md)|
|[Name](Office.ServerPolicy.Name.md)|
|[Parent](Office.ServerPolicy.Parent.md)|
|[Statement](Office.ServerPolicy.Statement.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
