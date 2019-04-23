---
title: Overview of the IRibbonUI Object
ms.prod: outlook
ms.assetid: ef273431-550f-4ff6-b964-79d05b09bea5
ms.date: 06/08/2017
localization_priority: Normal
---


# Overview of the IRibbonUI Object

An add-in can use the  [IRibbonUI](../../../api/Office.IRibbonUI.md) object to invalidate controls or menu items, and to update their content in the corresponding Microsoft Outlook user interface. The add-in specifies callback methods in the XML that [IRibbonExtensibility.GetCustomUI](../../../api/Office.IRibbonExtensibility.GetCustomUI.md) returns. These callback methods handle events for custom controls or custom menu items. 

When Outlook calls one of these methods, it passes an **IRibbonUI** object as a parameter to the callback method. The **IRibbonUI** object is scoped so that the add-in can only invalidate its own controls or menu items that use the object. The add-in cannot invalidate the controls or menu items that another add-in created.

 **IRibbonUI** exposes the following methods to customize the user interface in Outlook:


| **Method**| **Action**| **Description**|
|:-----|:-----|:-----|
| **[Invalidate()](../../../api/Office.IRibbonUI.Invalidate.md)**|Callback|Marks all of the custom controls or menu items in your add-in for update.|
| **[InvalidateControl(string controlID)](../../../api/Office.IRibbonUI.InvalidateControl.md)**|Callback|Marks a specific control or menu item that is defined by a  _controlID_ in your add-in for update.|
| **[ActivateTab](../../../api/Office.IRibbonUI.ActivateTab.md)**|Callback|Activates the specified custom tab on the Microsoft Office Fluent ribbon.|
| **[ActivateTabQ](../../../api/Office.IRibbonUI.ActivateTabQ.md)**|Callback|Activates the specified custom tab on the ribbon by using the fully qualified name of the tab.|

To minimize the impact on performance, use the  **InvalidateControl** method instead of the **Invalidate** method unless you actually need to invalidate all the custom controls or menu items that your add-in defines. Calling **Invalidate** invalidates all controls and menu items that your add-in defines, and callbacks occur on open explorers, inspectors, and menus.

## See also


 [Implementing the IRibbonExtensibility Interface](implementing-the-iribbonextensibility-interface.md)<br>
 [Office Fluent User Interface Extensibility for Outlook](office-fluent-user-interface-extensibility-for-outlook.md)<br>

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]