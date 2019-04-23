---
title: Cancel an event
ms.prod: access
ms.assetid: f91f4f8a-99fa-dca7-576a-11c76d6ddc93
ms.date: 09/25/2018
localization_priority: Normal
---


# Cancel an event

Under some circumstances, you may want to include code in an event procedure that cancels the associated event. For example, you may want to include code that cancels the **[Open](../../../api/Access.Form.Open.md)** event in an **Open** event procedure for a form, preventing the form from opening if certain conditions are not met.

You can cancel the following events:

- **ApplyFilter**
- **BeforeDelConfirm**
- **BeforeInsert**
- **BeforeRender**
- **BeforeUpdate**
- **CommandBeforeExecute**
- **DblClick**
- **Delete**
- **Dirty**
- **Exit**
- **Filter**
- **NoData**
- **Open**
- **Undo**
- **Unload**

You cancel an event by setting an event procedure's  _Cancel_ argument to **True**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]