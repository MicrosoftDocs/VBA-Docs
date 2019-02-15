---
title: RepaintObject, ShowAllRecords, Requery, and Refresh action/method comparison
keywords: vbaac10.chm5257635
f1_keywords:
- vbaac10.chm5257635
ms.prod: access
ms.assetid: ef1eec86-54d1-5b86-323f-48fb4f7d3897
ms.date: 02/16/2019
localization_priority: Normal
---


# RepaintObject, ShowAllRecords, Requery, and Refresh action/method comparison

The following table provides a brief comparison of the RepaintObject action, **[DoCmd.RepaintObject](../../../api/access.docmd.repaintobject.md)** method, **[Repaint](../../../api/access.form.repaint.md)** method, ShowAllRecords action, **[DoCmd.ShowAllRecords](../../../api/access.docmd.showallrecords.md)** method, Requery action, **[DoCmd.Requery](../../../api/access.docmd.requery.md)** method, **[Requery](../../../api/access.form.requery.md)** method, and **[Refresh](../../../api/access.form.refresh.md)** method.

|Action or method|Description|
|:---------------|:----------|
|RepaintObject action, **DoCmd.RepaintObject** method, **Repaint** method|Use the RepaintObject action, **RepaintObject** method or **Repaint** method to repaint controls in the specified object. They don't requery the database or display new records.|
|ShowAllRecords action, **DoCmd.ShowAllRecords** method|Use the ShowAllRecords action to requery and display the most recent records and remove any applied filters, which the Requery action doesn't do.|
|Requery action, **DoCmd.Requery** method|Use the Requery action or method to requery the source of the object or one of its controls. <br/><br/>The Requery action or method does one of the following: Reruns the query on which the control or object is based, displays any new or changed records, and removes any deleted records from the table on which the control or object is based.|
|**Refresh** method|Use the **Refresh** method to immediately update the records in the underlying record source for a specified form or datasheet to reflect changes made to the data by you and other users in a multiuser environment. The **Refresh** method shows only changes that have been made to the current set of records; it doesn't reflect new records or deleted records in the record source.|
|**Requery** method|Use the **Requery** method to update the data underlying a form or control to reflect records that are new to or have been deleted from the record source since it was last requeried. If you want to requery a control that isn't on the active object, you must use this method, not the Requery action or its corresponding **DoCmd.Requery** method.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]