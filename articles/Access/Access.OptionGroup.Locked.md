---
title: OptionGroup.Locked Property (Access)
keywords: vbaac10.chm10830
f1_keywords:
- vbaac10.chm10830
ms.prod: access
api_name:
- Access.OptionGroup.Locked
ms.assetid: 3d131a07-41cf-a21c-afad-623f01ed14ad
ms.date: 06/08/2017
---


# OptionGroup.Locked Property (Access)

The  **Locked** property specifies whether you can edit data in a control in Form view. Read/write **Boolean**.


## Syntax

 _expression_. **Locked**

 _expression_ A variable that represents an **OptionGroup** object.


## Remarks

The default setting of the  **Locked** property is **True**. This setting allows editing, adding, and deleting data.

Use the  **Locked** property to protect data in a field by making it read-only. For example, you might want a control to only display information without allowing editing, or you might want to lock a control until a specific condition is met.


## Example

The following example toggles the  **Enabled** property of a command button and the **Enabled** and **Locked** properties of a control, depending on the type of employee displayed in the current record. If the employee is a manager, then the SalaryDetails button is enabled and the PersonalInfo control is unlocked and enabled.


```vb
Sub Form_Current() 
 If Me!EmployeeType = "Manager" Then 
 Me!SalaryDetails.Enabled = True 
 Me!PersonalInfo.Enabled = True 
 Me!PersonalInfo.Locked = False 
 Else 
 Me!SalaryDetails.Enabled = False 
 Me!PersonalInfo.Enabled = False 
 Me!PersonalInfo.Locked = True 
 End If 
End Sub
```


## See also


#### Concepts


[OptionGroup Object](Access.OptionGroup.md)

