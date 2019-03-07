---
title: DoCmd.LockNavigationPane method (Access)
keywords: vbaac10.chm5853
f1_keywords:
- vbaac10.chm5853
ms.prod: access
api_name:
- Access.DoCmd.LockNavigationPane
ms.assetid: 64b44d9b-4cbd-182c-9bfb-89b4ca04dbf9
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.LockNavigationPane method (Access)

You can use the **LockNavigationPane** method to prevent users from deleting database objects that are displayed in the navigation pane.


## Syntax

_expression_.**LockNavigationPane** (_Lock_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Lock_|Required|**Variant**|Set to **True** to lock the navigation pane.|

## Remarks

Locking the navigation pane prevents the user from deleting database objects or cutting database objects to the clipboard. It does not prevent the user from performing any of the following operations:

- Copying database objects to the clipboard.
    
- Pasting database objects from the clipboard.
    
- Displaying or hiding the navigation pane.
    
- Selecting different navigation pane organization schemes.
    
- Showing or hiding sections of the navigation pane.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]