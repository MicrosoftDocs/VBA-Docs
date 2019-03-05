---
title: CustomControl.Custom property (Access)
keywords: vbaac10.chm12047
f1_keywords:
- vbaac10.chm12047
ms.prod: access
api_name:
- Access.CustomControl.Custom
ms.assetid: 9ce0028d-92a7-c113-c4c8-87caab8c4a5b
ms.date: 03/06/2019
localization_priority: Normal
---


# CustomControl.Custom property (Access)

Returns or sets a **String** representing the custom properties dialog box for an ActiveX control. Read/write.


## Syntax

_expression_.**Custom**

_expression_ A variable that represents a **[CustomControl](Access.CustomControl.md)** object.


## Remarks

Not all ActiveX controls provide a custom properties dialog box. To see whether a control provides this custom properties dialog box, look for the **Custom** property in the Microsoft Access property sheet for this control. If the list of properties contains the name **Custom**, the control provides the custom properties dialog box.

After you choose the **Custom** property box in the Microsoft Access property sheet, choose the **Build** button to the right of the property box to display the control's custom properties dialog box, often presented as a tabbed dialog box. Choose the tab that contains the interface for setting the properties that you want to set.

After you make changes on one tab, you can often apply those changes immediately by choosing the **Apply** button (if provided). You can choose other tabs to set other properties as needed. To approve all changes made in the custom properties dialog box, choose **OK**. To return to the Microsoft Access property sheet without changing any property settings, choose **Cancel**.

You can also view the custom properties dialog box by choosing the **Properties** subcommand of the ActiveX control **Object** command (for example, **Calendar Control Object**) on the **Edit** menu, or by choosing this same subcommand on the shortcut menu for the ActiveX control. In addition, some properties in the Microsoft Access property sheet for the ActiveX control, like the **GridFontColor** property of the Calendar control, have a **Build** button to the right of the property box. When you choose the **Build** button, the custom properties dialog box is displayed, with the appropriate tab selected (for example, **Colors**).


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]