---
title: SharedWorkspaceMembers.Add method (Office)
keywords: vbaof11.chm273003
f1_keywords:
- vbaof11.chm273003
ms.prod: office
api_name:
- Office.SharedWorkspaceMembers.Add
ms.assetid: 13d7c75d-a4d1-60ea-d689-c6886fb1e898
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceMembers.Add method (Office)

Adds a member to the list of members in a shared workspace site. Returns a **[SharedWorkspaceMember](Office.SharedWorkspaceMember.md)** object.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Add** (_Email_, _DomainName_, _DisplayName_, _Role_)

_expression_ Required. A variable that represents a **[SharedWorkspaceMembers](Office.SharedWorkspaceMembers.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Email_|Required|**String**|The new member's email address in the format user@domain.com. Raises an error if the user is not a valid candidate for membership in the shared workspace site.|
| _DomainName_|Required|**String**|The new member's Windows user name in the format domain\user.|
| _DisplayName_|Required|**String**|The display name to display for the new member.|
| _Role_|Optional|**String**|An optional role that determines the tasks that the new member can accomplish in the shared workspace site; for example, "Contributor." An invalid role name raises an error.|

## Example

The following example adds a new member to the members collection of the shared workspace site in the role of a site contributor.


```vb
    Dim swsMember As Office.SharedWorkspaceMember 
    Set swsMember = ActiveWorkbook.SharedWorkspace.Members.Add( _ 
        "user@domain.com", _ 
        "domain\user", _ 
        "New User", _ 
        "Contributor") 
    MsgBox "New member: " & swsMember.Name, _ 
        vbInformation + vbOKOnly, _ 
        "New Member in Shared Workspace)" 
    Set swsMember = Nothing 

```


## See also

- [SharedWorkspaceMembers object members](overview/Library-Reference/sharedworkspacemembers-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]