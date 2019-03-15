---
title: Report.OnCurrent property (Access)
keywords: vbaac10.chm13823
f1_keywords:
- vbaac10.chm13823
ms.prod: access
api_name:
- Access.Report.OnCurrent
ms.assetid: 593fdb6c-017a-986f-22ef-cc9e66aaaf01
ms.date: 03/15/2019
localization_priority: Normal
---


# Report.OnCurrent property (Access)

Sets or returns the value of the **OnCurrent** property on the report. Read/write **String**.


## Syntax

_expression_.**OnCurrent**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

If you want a procedure to run automatically every time you open a particular report, you set the form's **OnCurrent** property to [Event Procedure] and Access automatically stubs out a procedure for you called _Private Sub Report\_Current()_. 

The **OnCurrent** property allows you to programmatically determine the value of the form's **OnCurrent** property, or to programmatically set the form's **OnCurrent** property.

> [!NOTE] 
> The **[Current](Access.Report.Current.md)** event fires when you run (open) a report.

If you set the form's **OnCurrent** property in the UI, it gets its value based on your selection in the Choose Builder window, which appears when you choose the **...** button next to the **On Current** box in the report's Properties window.

- If you choose Expression Builder, the value will be =_expression_, where _expression_ is the expression from the Expression Builder window.
    
- If you choose Macro Builder, the value is the name of the macro. 
    
- If you choose Code Builder, the value will be [Event Procedure].
    

## Example

The following code example demonstrates how to set a report's **OnCurrent** property.

```vb

Private Sub Report_Load()

        Me.OnCurrent = "[Event Procedure]"

End Sub
		
```

<br/>

The event procedure **Report_Current()** is automatically called when the **Current** event is fired. This procedure simply collects the values of two of the report's text boxes and sends them to another procedure for processing.

```vb

Private Sub Report_Current()

        ' Declare variables to store price and available credit.
        Dim curPrice As Currency
        Dim curCreditAvail As Currency

        ' Assign variables from current values in text boxes on the Report.
        curPrice = txtValue1
        curCreditAvail = txtValue2

        ' Call VerifyCreditAvail procedure.
        VerifyCreditAvail curPrice, curCreditAvail

End Sub
		
```

<br/>

The following code example simply processes the two values passed to it.

```vb
Sub VerifyCreditAvail(curTotalPrice As Currency, curAvailCredit As Currency)
    ' Inform the user if there is not enough credit available for the purchase.
    If curTotalPrice > curAvailCredit Then
        MsgBox "You do not have enough credit available for this purchase."
    End If
End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]