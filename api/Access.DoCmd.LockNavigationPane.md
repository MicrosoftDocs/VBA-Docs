---
title: DoCmd.LockNavigationPane method (Access)
keywords: vbaac10.chm5853
f1_keywords:
- vbaac10.chm5853
ms.prod: access
api_name:
- Access.DoCmd.LockNavigationPane
ms.assetid: 64b44d9b-4cbd-182c-9bfb-89b4ca04dbf9
ms.date: 06/08/2017
localization_priority: Normal
---


# DoCmd.LockNavigationPane method (Access)

You can use the  **LockNavigationPane** action to prevent users from deleting database objects that are displayed in the Navigation Pane.


## Syntax

_expression_. `LockNavigationPane`( ` _Lock_` )

_expression_ A variable that represents a [DoCmd](Access.DoCmd.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Lock_|Required|**Variant**|Set to  **True** to lock the Navigation Pane.|

## Remarks

Locking the Navigation Pane prevents the user from deleting database objects or cutting database objects to the clipboard. It does not prevent the user from performing any of the following operations:


- Copying database objects to the clipboard
    
- Pasting database objects from the clipboard
    
- Displaying or hiding the Navigation Pane
    
- Selecting different Navigation Pane organization schemes
    
- Showing or hiding sections of the Navigation Pane
    

## See also


[DoCmd Object](Access.DoCmd.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]