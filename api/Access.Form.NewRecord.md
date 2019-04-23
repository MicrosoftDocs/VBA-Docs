---
title: Form.NewRecord property (Access)
keywords: vbaac10.chm13491
f1_keywords:
- vbaac10.chm13491
ms.prod: access
api_name:
- Access.Form.NewRecord
ms.assetid: 9e30b019-1c1d-31eb-cc8d-cab030861ddc
ms.date: 03/14/2019
localization_priority: Normal
---


# Form.NewRecord property (Access)

You can use the **NewRecord** property to determine whether the current record is a new record. Read-only **Integer**.


## Syntax

_expression_.**NewRecord**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The **NewRecord** property is read-only in Form view and Datasheet view. It isn't available in Design view.

When a user has moved to a new record, the **NewRecord** property setting will be **True** whether or not the user has started to edit the record.


## Example

The following example shows how to use the **NewRecord** property to determine if the current record is a new record. The **NewRecordMark** procedure sets the current record to the variable _intnewrec_. If the record is new, a message is displayed notifying the user of this. You could run this procedure when the **Current** event for a form occurs.

```vb
Sub NewRecordMark(frm As Form) 
 Dim intnewrec As Integer 
 
 intnewrec = frm.NewRecord 
 If intnewrec = True Then 
 MsgBox "You're in a new record." _ 
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
