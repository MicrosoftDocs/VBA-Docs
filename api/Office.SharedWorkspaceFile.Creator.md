---
title: SharedWorkspaceFile.Creator property (Office)
ms.prod: office
api_name:
- Office.SharedWorkspaceFile.Creator
ms.assetid: beae3af9-e256-65ba-3814-8b8944910e2a
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceFile.Creator property (Office)

Gets a 32-bit integer that indicates the application in which the **SharedWorkspaceFile** object was created. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[SharedWorkspaceFile](Office.SharedWorkspaceFile.md)** object.


## Return value

Long


## Remarks

As an example, if the object was created in Microsoft Word, this property returns 1297307460, which represents the string "MSWD"; in Microsoft Excel, this property returns 1480803660. This value can also be represented by the constant **wdCreatorCode** in Word or **xlCreatorCode** in Excel. 

The **Creator** property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with the Microsoft Office Macintosh Edition.

The **Creator** property always returns the numeric identifier for the active application, just as the **Application** property always returns the name of the active application in string form. Use the **CreatedBy** property of the **SharedWorkspaceFile** object to return the name of the individual who created the object. Use document properties to return information about the authors of Office documents.


## See also

- [SharedWorkspaceFile object members](overview/Library-Reference/sharedworkspacefile-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]