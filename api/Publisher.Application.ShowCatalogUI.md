---
title: Application.ShowCatalogUI event (Publisher)
keywords: vbapb10.chm268435493
f1_keywords:
- vbapb10.chm268435493
ms.prod: publisher
api_name:
- Publisher.Application.ShowCatalogUI
ms.assetid: 8a5a3798-4b95-d77f-70f6-d69dd9dc8f99
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.ShowCatalogUI event (Publisher)

Fires when the catalog of publication wizards is displayed in the Microsoft Publisher user interface.


## Syntax

_expression_.**ShowCatalogUI**

_expression_ An expression that returns an **[Application](Publisher.Application.md)** object.


## Remarks

You can use the **[ShowWizardCatalog](Publisher.Application.ShowWizardCatalog.md)** method to display the wizard catalog in the user interface.

The **ShowCatalogUI** event does not fire when the publication catalog is displayed when Publisher first starts. To determine if the catalog is displayed at that time, you can use the **[WizardCatalogVisible](Publisher.Application.WizardCatalogVisible.md)** property.

For more information about using events with the **Application** object, see [Using events with the Application object](../publisher/Concepts/using-events-with-the-application-object-publisher.md).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to handle the **ShowCatalogUI** event. It displays a message notifying the user that the catalog UI was displayed.

```vb
Private Sub pubApplication_ShowCatalogUI() 
 MsgBox "The Wizard Catalog is displayed." 
End Sub
```

<br/>

For this event to occur, you must place the following line of code in the General Declarations section of your module.

```vb
Private WithEvents pubApplication As Application
```

<br/>

You then must run the following initialization procedure.

```vb
Public Sub Initialize_pubApplication() 
 Set pubApplication = Publisher.Application 
End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]