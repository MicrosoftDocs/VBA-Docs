---
title: Report.NoData event (Access)
keywords: vbaac10.chm13881
f1_keywords:
- vbaac10.chm13881
ms.prod: access
api_name:
- Access.Report.NoData
ms.assetid: fa5f22b1-3695-bd16-2ca3-b2a1cc1f1d94
ms.date: 03/08/2019
localization_priority: Normal
---


# Report.NoData event (Access)

The **NoData** event occurs after Microsoft Access formats a report for printing that has no data (the report is bound to an empty recordset), but before the report is printed. You can use this event to cancel printing of a blank report.

## Syntax

_expression_.**NoData** (_Cancel_)

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|The setting determines whether to print the report. Setting the _Cancel_ argument to **True** (1) prevents the report from printing. You can also use the **CancelEvent** method of the **DoCmd** object to cancel printing the report.|

## Remarks

To run a macro or event procedure when this event occurs, set the **[OnNoData](Access.Report.OnNoData.md)** property to the name of the macro or to [Event Procedure].

If the report isn't bound to a table or query (by using the report's **[RecordSource](Access.Report.RecordSource.md)** property), the **NoData** event doesn't occur.

This event occurs after the **Format** events for the report, but before the first **Print** event.

This event doesn't occur for subreports. If you want to hide controls on a subreport when the subreport has no data, so that the controls don't print in this case, you can use the **HasData** property in a macro or event procedure that runs when the **Format** or **Print** event occurs.

The **NoData** event occurs before the first **Page** event for the report.


## Example

The following example shows how to cancel printing a report when it has no data. A message box notifying the user that the printing has been canceled is also displayed. 

To try this example, add the following event procedure to a report. Try running the report when it contains no data. 

```vb
Private Sub Report_NoData(Cancel As Integer) 
    MsgBox "The report has no data." & _ 
         chr(13) & "Printing is canceled. " & _ 
         chr(13) & "Check the data source for the " & _ 
         chr(13) & "report. Make sure you entered " & _ 
         chr(13) & "the correct criteria (for " & _ 
         chr(13) & "example, a valid range of " & _ 
         chr(13) & "dates),." vbOKOnly + vbInformation 
    Cancel = True 
End Sub 
```

<br/>

The following example shows how to use the **NoData** event of a report to prevent the report from opening when there is no data to be displayed.

```vb
Private Sub Report_NoData(Cancel As Integer)

    'Add code here that will be executed if no data
    'was returned by the Report's RecordSource
    MsgBox "No customers ordered this product this month. " & _
        "The report will now close."
    Cancel = True

End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]