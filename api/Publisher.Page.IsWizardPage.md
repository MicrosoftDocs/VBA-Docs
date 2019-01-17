---
title: Page.IsWizardPage Property (Publisher)
keywords: vbapb10.chm393271
f1_keywords:
- vbapb10.chm393271
ms.prod: publisher
api_name:
- Publisher.Page.IsWizardPage
ms.assetid: 09c1352d-6760-ad54-aa95-211727c968b3
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.IsWizardPage Property (Publisher)

Returns  **True** if the specified page is a Microsoft Publisher wizard page. Read-only **Boolean**.


## Syntax

 _expression_. **IsWizardPage**

 _expression_ A variable that represents an  **Page** object.


## Return value

Boolean


## Remarks

Wizard pages are special page types for certain types of Publisher wizards (such as Newsletters, Catalogs, and Web Wizards) that can be inserted into a publication.

Use the  **[Wizard](Publisher.Page.Wizard.md)** property of the **[Page](Publisher.Page.md)** object to access the wizard for the specified page.


## Example

The following example tests to determine whether the specified page is a wizard page. If it is, certain wizard properties are returned.


```vb
 With ActiveDocument.Pages(1) 
 If .IsWizardPage = True Then 
 
 With .Wizard 
 Debug.Print .Name 
 Debug.Print .Properties(1).Name 
 Debug.Print .Properties(1).CurrentValueId 
 End With 
 
 End If 
 End With
```


