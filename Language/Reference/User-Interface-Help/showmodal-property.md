---
title: ShowModal property (Visual Basic for Applications)
keywords: vblr6.chm1248574
f1_keywords:
- vblr6.chm1248574
ms.prod: office
api_name:
- Office.ShowModal
ms.assetid: 710c7bc7-ce50-057f-680e-e2be157d0dac
ms.date: 12/19/2018
localization_priority: Normal
---


# ShowModal property

Sets a **[UserForm](userform-object.md)** to be modal or modeless in its display. Read-only at [run time](../../Glossary/vbe-glossary.md#run-time).

## Remarks

The settings for the **ShowModal** property are:

|Setting|Description|
|:-----|:-----|
|**True**|(Default) The **UserForm** is modal.|
|**False**|The **UserForm** is modeless.|


## Remarks

When a **UserForm** is modal, the user must supply information or close the **UserForm** before using any other part of the application. No subsequent code is executed until the **UserForm** is hidden or unloaded. Although other forms in the application are disabled when a **UserForm** is displayed, other applications are not.

When the **UserForm** is modeless, the user can view other forms or windows without closing the **UserForm**.
Modeless forms do not appear in the task bar and are not in the window tab order.

> [!NOTE] 
> If you attempt to open a **UserForm** that has a **ShowModal** property in Microsoft Office 97, you get a run-time error because Office 97 only displays modal **UserForms**. Office 97 ignores the **ShowModal** property and displays the **UserForm** modally.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
