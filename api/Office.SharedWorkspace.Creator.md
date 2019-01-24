---
title: SharedWorkspace.Creator property (Office)
ms.prod: office
api_name:
- Office.SharedWorkspace.Creator
ms.assetid: 167fdd22-50ab-9b27-f594-27c38d88a4a9
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspace.Creator property (Office)

Gets a 32-bit integer that indicates the application in which the **SharedWorkspace** object was created. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[SharedWorkspace](Office.SharedWorkspace.md)** object.


## Return value

Long


## Remarks

As an example, if the object was created in Microsoft Word, this property returns 1297307460, which represents the string "MSWD"; in Microsoft Excel, this property returns 1480803660. This value can also be represented by the constant **wdCreatorCode** in Word or **xlCreatorCode** in Excel. 

The **Creator** property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with the Microsoft Office Macintosh Edition.


## See also

- [SharedWorkspace object members](overview/Library-Reference/sharedworkspace-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]