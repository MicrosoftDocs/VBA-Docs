---
title: SharedWorkspaceTask.Creator Property (Office)
ms.prod: office
api_name:
- Office.SharedWorkspaceTask.Creator
ms.assetid: 414bba79-c72b-04f6-9784-3fb959b593e1
ms.date: 06/08/2017
---


# SharedWorkspaceTask.Creator Property (Office)

Gets a 32-bit integer that indicates the application in which the  **SharedWorkspaceTask** object was created. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. `Creator`

 _expression_ A variable that represents a [SharedWorkspaceTask](./Office.SharedWorkspaceTask.md) object.


### Return value

Long


## Remarks

As an example, if the object was created in Microsoft Word, this property returns 1297307460, which represents the string "MSWD"; in Microsoft Excel, this property returns 1480803660. This value can also be represented by the constant wdCreatorCode in Word, or xlCreatorCode in Excel. The  **Creator** property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.

The  **Creator** property always returns the numeric identifier for the active application, just as the **Application** property always returns the name of the active applicatin in string form. Use the **CreatedBy** property of the **SharedWorkspaceTask** object to return the name of the individual who created the object. Use document properties to return information about the authors of Office documents.


## Example

This example displays a message about the creator of "myObject" variable.


```vb
Set myObject = ActiveDocument 
If myObject.Creator = wdCreatorCode Then 
    MsgBox "This is a Microsoft Word object" 
Else 
    MsgBox "This is not a Microsoft Word object" 
End If 

```


## See also


[SharedWorkspaceTask Object](Office.SharedWorkspaceTask.md)



[SharedWorkspaceTask Object Members](./overview/Library-Reference/sharedworkspacetask-members-office.md)

