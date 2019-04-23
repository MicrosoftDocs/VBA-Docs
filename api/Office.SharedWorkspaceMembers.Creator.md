---
title: SharedWorkspaceMembers.Creator property (Office)
ms.prod: office
api_name:
- Office.SharedWorkspaceMembers.Creator
ms.assetid: 0b43590b-67f2-68a6-3117-4972754aa7c8
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceMembers.Creator property (Office)

Gets a 32-bit integer that indicates the application in which the **SharedWorkspaceMembers** object was created. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[SharedWorkspaceMembers](Office.SharedWorkspaceMembers.md)** object.


## Return value

Long


## Remarks

As an example, if the object was created in Microsoft Word, this property returns 1297307460, which represents the string "MSWD"; in Microsoft Excel, this property returns 1480803660. This value can also be represented by the constant **wdCreatorCode** in Word or **xlCreatorCode** in Excel. 

The **Creator** property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with the Microsoft Office Macintosh Edition.

The **Creator** property always returns the numeric identifier for the active application, just as the **Application** property always returns the name of the active application in string form. Use the **CreatedBy** property of the **SharedWorkspaceFile**, **SharedWorkspaceLink**, and **SharedWorkspaceTask** objects to return the name of the individual who created those objects. Use document properties to return information about the authors of Office documents.


## See also

- [SharedWorkspaceMembers object members](overview/Library-Reference/sharedworkspacemembers-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]