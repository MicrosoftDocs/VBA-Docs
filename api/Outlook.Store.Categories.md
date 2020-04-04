---
title: Store.Categories property (Outlook)
keywords: vbaol11.chm3512
f1_keywords:
- vbaol11.chm3512
ms.prod: outlook
api_name:
- Outlook.Store.Categories
ms.assetid: 597678d0-51f6-45d7-a98a-063344bbcff7
ms.date: 06/08/2017
localization_priority: Normal
---


# Store.Categories property (Outlook)

Returns a **[Categories](Outlook.Categories.md)** collection that represents all of the categories that are defined for the **[Store](Outlook.Store.md)**. Read-only.


## Syntax

_expression_. `Categories`

_expression_ A variable that represents a [Store](Outlook.Store.md) object.


## Remarks

This property is similar to the  **[Categories](Outlook.NameSpace.Categories.md)** property of the **[NameSpace](Outlook.NameSpace.md)** object, except that the **Store.Categories** property applies to a session profile that specifies one or more accounts and **Store.Categories** specifies the categories for the store that an account is associated with, whereas **NameSpace.Categories** applies to a session profile that defines only one account and the **NameSpace.Categories** property specifies the Master Category List for that session.

For certain secondary delivery stores such as an IMAP store, the  **Categories** property returns the **Categories** collection for the primary store. IMAP stores do not actually support a separate categories collection.


## Example

The following code sample in Microsoft Visual Basic for Applications (VBA) displays the name of each  **[Category](Outlook.Category.md)** object that is contained in the **Categories** collection associated with each **Store** object in the **[Stores](Outlook.Stores.md)** collection for the session.


```vb
Sub EnumerateCategoriesForStores() 
 
 Dim oStores As Outlook.Stores 
 
 Dim oStore As Outlook.Store 
 
 Dim oCategories As Outlook.Categories 
 
 Dim oCategory As Outlook.Category 
 
 Set oStores = Application.Session.Stores 
 
 For Each oStore In oStores 
 
 Debug.Print oStore.DisplayName 
 
 Debug.Print "--------------Categories-----------------" 
 
 Set oCategories = oStore.Categories 
 
 For Each oCategory In oCategories 
 
 Debug.Print Chr(9) & oCategory.Name 
 
 Next 
 
 Debug.Print "" 
 
 Next 
 
End Sub
```


## See also


[Store Object](Outlook.Store.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]