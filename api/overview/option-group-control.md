---
title: Option Group Control
keywords: vbaac10.chm13398
f1_keywords:
- vbaac10.chm13398
ms.prod: access
ms.assetid: a67b22b7-d3a8-c9c6-cb1b-a6d544b2fefe
ms.date: 06/08/2017
localization_priority: Normal
---


# Option Group Control

**Applies to:** Access 2013 | Access 2016

An option group on a form or report displays a limited set of alternatives. An option group makes selecting a value easy since you can just click the value you want. Only one option in an option group can be selected at a time.

An option group consists of a group frame and a set of check boxes, toggle buttons, or option buttons.

## Remarks

If an option group is bound to a field, only the group frame itself is bound to the field, not the check boxes, toggle buttons, or option buttons inside the frame. Instead of setting the  **ControlSource** property for each control in the option group, you set the **OptionValue** property of each check box, toggle button, or option button to a number that's meaningful for the field to which the group frame is bound. When you select an option in an option group, Microsoft Access sets the value of the field to which the option group is bound to the value of the selected option's **OptionValue** property.

> [!NOTE] 
> The  **OptionValue** property is set to a number because the value of an option group can only be a number, not text. Microsoft Access stores this number in the underlying table. In the preceding example, if you want to display the name of the shipper instead of a number in the Orders table, you can create a separate table called Shippers that stores shipper names, and then make the ShipVia field in the Orders table a Lookup field that looks up data in the Shippers table.

An option group can also be set to an expression, or it can be unbound. You can use an unbound option group in a custom dialog box to accept user input and then carry out an action based on that input.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Search for specific Access error codes on Bing](https://www.bing.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access wiki on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]