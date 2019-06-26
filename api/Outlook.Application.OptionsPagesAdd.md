---
title: Application.OptionsPagesAdd event (Outlook)
keywords: vbaol11.chm432
f1_keywords:
- vbaol11.chm432
ms.prod: outlook
api_name:
- Outlook.Application.OptionsPagesAdd
ms.assetid: aa13cd97-de96-00f8-a532-ca8ee9b00343
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.OptionsPagesAdd event (Outlook)

Occurs whenever the user clicks the  **Add-in Options** button on the **Add-ins** tab of the Outlook **Options** dialog box.


## Syntax

_expression_. `OptionsPagesAdd`( `_Pages_` , `_Folder_` )

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pages_|Required| **[PropertyPages](Outlook.PropertyPages.md)**|The collection of property pages that have been added to the dialog box. This collection includes only custom property pages. It does not include standard Microsoft Outlook property pages.|
| _Folder_|Required| **PropertyPages**|This argument is only used with the  **[Folder](Outlook.Folder.md)** object. The **Folder** object for which the **Properties** dialog box is being opened.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).

Your program handles this event to add a custom property page. The property page will be added to the  **Options** dialog box. When the event fires, the **PropertyPages** collection object identified by _Pages_ contains the property pages that have been added prior to the event handler being called. To add your property page to the collection, use the **[Add](Outlook.PropertyPages.Add.md)** method of the **PropertyPages** collection before exiting the event handler.


## Example

This Microsoft Visual Basic for Applications (VBA) example adds a new property page to the Outlook  **Options** dialog box. The sample code must be placed in a class module of a Component Object Model (COM) add-in. For information about COM add-ins, see [Customizing Outlook using COM add-ins](../outlook/Concepts/Getting-Started/customizing-outlook-using-com-add-ins.md).


```vb
Implements IDTExtensibility2 
Private WithEvents OutlApp As Outlook.Application 
 
Private Sub IDTExtensibility2_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant) 
 Set OutlApp = Outlook.Application 
End Sub 
 
Private Sub OutlApp_OptionsPagesAdd(ByVal Pages As Outlook.PropertyPages) 
 Pages.Add "PPE.SimplePage", "Simple Page" 
 'PPE.SimplePage is a ProgID of the registered ActiveX Control - the property page that is to be displayed in the COM add-in 
End Sub
```


## See also


[Application Object](Outlook.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]