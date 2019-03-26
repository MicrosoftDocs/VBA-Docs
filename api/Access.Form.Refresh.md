---
title: Form.Refresh method (Access)
keywords: vbaac10.chm13504
f1_keywords:
- vbaac10.chm13504
ms.prod: access
api_name:
- Access.Form.Refresh
ms.assetid: e7a15c34-d3ec-184f-8d03-3e264fcc60d0
ms.date: 03/09/2019
localization_priority: Priority
---


# Form.Refresh method (Access)

The **Refresh** method immediately updates the records in the underlying record source for a specified form or datasheet to reflect changes made to the data by you and other users in a multiuser environment.


## Syntax

_expression_.**Refresh**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Return value

Nothing


## Remarks

Using the **Refresh** method is equivalent to choosing **Refresh** on the **Home** tab.

Microsoft Access refreshes records automatically, based on the **Refresh Interval** setting on the **Advanced** tab of the **Access Options** dialog box, available by choosing the Microsoft Office button, and then choosing **Access Options**. ODBC data sources are refreshed based on the **ODBC Refresh Interval** setting on the **Advanced** tab of the **Access Options** dialog box. You can use the **Refresh** method to view changes that have been made to the current set of records in a form or datasheet since the record source underlying the form or datasheet was last refreshed.

In an Access database, the **Refresh** method shows only changes made to records in the current set. Because the **Refresh** method doesn't actually requery the database, the current set won't include records that have been added or exclude records that have been deleted since the database was last requeried, nor will it exclude records that no longer satisfy the criteria of the query or filter. To requery the database, use the **[Requery](Access.Form.Requery.md)** method. When the record source for a form is requeried, the current set of records will accurately reflect all data in the record source.

In an Access project (.adp), the **Refresh** method requeries the database and displays any new or changed records or removes deleted records from the table on which the form is based. The form is also updated to display records based on any changes to the **[Filter](Access.Form.Filter(property).md)** property of the form.

> [!NOTE] 
> - It's often faster to refresh a form or datasheet than to requery it. This is especially true if the initial query was slow to run.
> - Don't confuse the **Refresh** method with the **[Repaint](Access.Form.Repaint.md)** method, which repaints the screen with any pending visual changes.

## Example

The following example uses the **Refresh** method to update the records in the underlying record source for the **Customers** form whenever the form receives the focus.

```vb
Private Sub Form_Activate() 
    Me.Refresh 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
