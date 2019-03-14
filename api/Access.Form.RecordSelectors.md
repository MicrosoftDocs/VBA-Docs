---
title: Form.RecordSelectors property (Access)
keywords: vbaac10.chm13364
f1_keywords:
- vbaac10.chm13364
ms.prod: access
api_name:
- Access.Form.RecordSelectors
ms.assetid: 7700f0c5-621f-5145-57be-777d53228379
ms.date: 03/14/2019
localization_priority: Normal
---


# Form.RecordSelectors property (Access)

You can use the **RecordSelectors** property to specify whether a form displays record selectors in Form view. Read/write **Boolean**.


## Syntax

_expression_.**RecordSelectors**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The default value is **True**.

You can use this property to remove record selectors when you create or use a form as a custom dialog box or a palette. You can also use this property for forms whose **[DefaultView](Access.Form.DefaultView.md)** property is set to Single Form.

The record selector displays the unsaved record indicator when a record is being edited. When the **RecordSelectors** property is set to No and the **[RecordLocks](Access.Form.RecordLocks.md)** property is set to Edited Record (record locking is "pessimistic"—only one person can edit a record at a time), there is no visual clue that the record is locked.


## Example

The following example specifies that no record has a record selector in the **Employees** form.

```vb
Forms("Employees").RecordSelectors = False
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]