---
title: Overview of the IRibbonControl Object
ms.prod: outlook
ms.assetid: 32a0ae0b-26d9-673b-d609-b86696538435
ms.date: 06/08/2019
localization_priority: Normal
---


# Overview of the IRibbonControl Object

The [IRibbonControl](../../../api/Office.IRibbonControl.md) object is passed in most of the callbacks that are available for controls in the ribbon or Microsoft Office Backstage view, as well as the customizable menu items in Microsoft Outlook. The object is especially useful for Outlook developers because it provides an [IRibbonControl.Context](../../../api/Office.IRibbonControl.Context.md) property that returns the related Outlook object to which the customization is applied and is about to be displayed. 

For example, the **Context** property returns the [Explorer](../../../api/Outlook.Explorer.md) object if you customize the ribbon in an explorer, and returns the [Store](../../../api/Outlook.Store.md) object if you customize the shortcut menu for a store folder.

 **IRibbonControl** exposes the following properties.


| **Property**| **Type**| **Description**|
|:-----|:-----|:-----|
| **Context**| **Object**| Returns an object that represents the window in which the custom ribbon is about to be displayed, or the related object to which the menu customization is applied and is about to be displayed. Read-only.|
| **[Id](../../../api/Office.IRibbonControl.Id.md)**| **String**|Returns a string that represents the **Id** attribute for the control or custom menu item. Read-only.|
| **[Tag](../../../api/Office.IRibbonControl.Tag.md)**| **String**|Returns a string that represents the **Tag** attribute for the control or custom menu item. Read-only.|

When you write managed code, try to cast the object represented by **IRibbonControl.Context** to the corresponding Outlook object. For example, if you customize a ribbon in an inspector, cast the [Inspector](../../../api/Outlook.Inspector.md) object. Then, if the cast succeeds, you can compare the **Inspector** object that is returned by **IRibbonControl.Context** to other inspector windows that are open. To determine the underlying item that is displayed in an inspector window, examine [Inspector.CurrentItem](../../../api/Outlook.Inspector.CurrentItem.md). Because **CurrentItem** is an **Object** type, your code must cast the object to an appropriate item type such as [MailItem](../../../api/Outlook.MailItem.md) or [ContactItem](../../../api/Outlook.ContactItem.md).

## See also


 [Office Fluent User Interface Extensibility for Outlook](office-fluent-user-interface-extensibility-for-outlook.md)<br>
 [Implementing the IRibbonExtensibility Interface](implementing-the-iribbonextensibility-interface.md)<br>

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
