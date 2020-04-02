---
title: Office Fluent User Interface Extensibility for Outlook
ms.prod: outlook
ms.assetid: 8496c52e-1f9d-16ef-2fd8-c1bca1a96816
ms.date: 06/08/2019
localization_priority: Normal
---


# Office Fluent User Interface Extensibility for Outlook

Microsoft Office Fluent user interface (UI) extensibility refers to the mechanism that supports the programmatic customization of the Office Fluent UI in Microsoft Office applications. Microsoft Outlook expands user interface extensibility beyond the ribbons in the explorer and inspector windows to include other components of the Outlook UI, such as the following:


- Microsoft Office Backstage view
    
- Contextual tabs
    
- New items menus
    
- Shortcut menus
    



 Add-ins implement the **[IRibbonExtensibility](../../../api/Office.IRibbonExtensibility.md)** interface to extend the Outlook UI. To customize a part of the UI, specify your customizations in an XML markup file that complies with the schema definition for Office Fluent UI extensibility. Office calls the **[IRibbonExtensibility.GetCustomUI](../../../api/Office.IRibbonExtensibility.GetCustomUI.md)** method, specifying a ribbon ID to load the XML that describes your customizations for the part of the Outlook UI that matches the ribbon ID. As a result of the XML markup, your add-in runs callback procedures that execute the custom actions that are associated with commands in the custom UI.
 
Unlike other Office applications such as Microsoft Word or Microsoft Excel, you cannot customize the ribbon by using Visual Basic for Applications in Outlook. To programmatically customize the UI in Outlook, you must write an add-in. You can update an existing add-in or write an add-in that only targets Outlook. The add-in can be native or managed. Outlook does not support the customization of the ribbon by using Microsoft Office Open XML Format Files. 

For more information and examples of different ways to customize the Outlook UI, see [Extending the User Interface in Outlook 2010](https://msdn.microsoft.com/library/ee692172%28office.14%29.aspx) on the MSDN Web site.

## See also


 [Customizing Shortcut Menus](customizing-shortcut-menus.md)<br>
 [Updating Earlier Code for CommandBars](updating-earlier-code-for-commandbars.md)<br>
 [Implementing the IRibbonExtensibility Interface](implementing-the-iribbonextensibility-interface.md)<br>
 [Overview of Customizing the Ribbon](overview-of-customizing-the-ribbon.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]