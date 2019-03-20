---
title: Report.ServerFilter property (Access)
keywords: vbaac10.chm13694
f1_keywords:
- vbaac10.chm13694
ms.prod: access
api_name:
- Access.Report.ServerFilter
ms.assetid: e73ad797-8c76-705f-080b-2d0f3423cb39
ms.date: 03/20/2019
localization_priority: Normal
---


# Report.ServerFilter property (Access)

You can use the **ServerFilter** property to specify a subset of records to be displayed when a server filter is applied to a report within a Microsoft Access project (.adp) or database. Read/write **String**.


## Syntax

_expression_.**ServerFilter**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

The **ServerFilter** property is a string expression consisting of a WHERE clause without the WHERE keyword. For example, the following Visual Basic code defines and applies a filter to show only customers from the USA.

```vb
Me.ServerFilter = "Country = 'USA'" 
Me.Refresh
```

To set the **ServerFilter** property, you must first either:

- Set the property value in the form's property sheet.
    
- Set the property in Visual Basic by typing the following code:
    
  ```vb
    Forms(0).ServerFilter = "fieldname = value "
  ```

> [!NOTE] 
> Setting the **ServerFilter** property has no effect on the ADO **Filter** property.

You can use the **ServerFilter** property to save a filter and apply it at a later time. Filters are saved with the objects in which they are created. They are automatically loaded when the object is opened, but they aren't automatically applied.

To apply a saved filter to a form, you can choose **Apply Server Filter** on the toolbar, choose **Apply Filter/Sort** on the **Records** menu, or use a macro or Visual Basic to set the **ServerFilterByForm** property to **True**.

The **Apply Server Filter** button indicates the state of the **ServerFilter** and **ServerFilterByForm** properties. The button remains disabled until there is a filter to apply. If an existing filter is currently applied, the **Apply Server Filter** button appears pressed in.

To apply a filter automatically when a form is opened, specify in the **OnOpen** event property setting of the form either a macro that uses the ApplyFilter action or an event procedure that uses the **ApplyFilter** method of the **DoCmd** object. In either case, the form opens in the Server Filter By Form window.

You can only remove a server filter by using Visual Basic to set the **ServerFilterByForm** property to **False** or clear all filter criteria in the Server Filter By Form window, and then choose **Apply Server Filter**.

When the **ServerFilter** property is set in form Design view, Access does not attempt to validate the SQL expression. If the SQL expression is invalid, an error occurs when the filter is applied.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]