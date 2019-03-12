---
title: AccessObject.IsLoaded property (Access)
keywords: vbaac10.chm12750
f1_keywords:
- vbaac10.chm12750
ms.prod: access
api_name:
- Access.AccessObject.IsLoaded
ms.assetid: 5e68398c-8a95-f3e1-87ec-e2d637f34429
ms.date: 02/01/2019
localization_priority: Normal
---


# AccessObject.IsLoaded property (Access)

You can use the **IsLoaded** property to determine if an **AccessObject** object is currently loaded. Read-only **Boolean**.


## Syntax

_expression_.**IsLoaded**

_expression_ A variable that represents an **[AccessObject](Access.AccessObject.md)** object.


## Remarks

The **IsLoaded** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Yes|**True**|The specified **AccessObject** is loaded.|
|No|**False**|The specified **AccessObject** is not loaded.|

## Example

The following example shows how to prevent a user from opening a particular form directly from the navigation pane.

```vb
'Don't let this form be opened from the Navigator
If Not CurrentProject.AllForms(cFormUsage).IsLoaded Then
    MsgBox "This form cannot be opened from the navigation pane.", _
        vbInformation + vbOKOnly, "Invalid form usage"
    Cancel = True
    Exit Sub
End If
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
