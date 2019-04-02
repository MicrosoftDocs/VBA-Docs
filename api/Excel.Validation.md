---
title: Validation object (Excel)
keywords: vbaxl10.chm531072
f1_keywords:
- vbaxl10.chm531072
ms.prod: excel
api_name:
- Excel.Validation
ms.assetid: 59d29d1e-92d3-373e-04d0-0d7fe97e1878
ms.date: 04/03/2019
localization_priority: Normal
---


# Validation object (Excel)

Represents data validation for a worksheet range.


## Example

Use the **[Validation](Excel.Range.Validation.md)** property of the **Range** object to return the **Validation** object. The following example changes the data validation for cell E5.

```vb
Range("e5").Validation _ 
 .Modify xlValidateList, xlValidAlertStop, "=$A$1:$A$10"
```

<br/>

Use the **Add** method to add data validation to a range and create a new **Validation** object. The following example adds data validation to cell E5.

```vb
With Range("e5").Validation 
 .Add Type:=xlValidateWholeNumber, _ 
 AlertStyle:=xlValidAlertInformation, _ 
 Minimum:="5", Maximum:="10" 
 .InputTitle = "Integers" 
 .ErrorTitle = "Integers" 
 .InputMessage = "Enter an integer from five to ten" 
 .ErrorMessage = "You must enter a number from five to ten" 
End With 

```


## Methods

- [Add](Excel.Validation.Add.md)
- [Delete](Excel.Validation.Delete.md)
- [Modify](Excel.Validation.Modify.md)

## Properties

- [AlertStyle](Excel.Validation.AlertStyle.md)
- [Application](Excel.Validation.Application.md)
- [Creator](Excel.Validation.Creator.md)
- [ErrorMessage](Excel.Validation.ErrorMessage.md)
- [ErrorTitle](Excel.Validation.ErrorTitle.md)
- [Formula1](Excel.Validation.Formula1.md)
- [Formula2](Excel.Validation.Formula2.md)
- [IgnoreBlank](Excel.Validation.IgnoreBlank.md)
- [IMEMode](Excel.Validation.IMEMode.md)
- [InCellDropdown](Excel.Validation.InCellDropdown.md)
- [InputMessage](Excel.Validation.InputMessage.md)
- [InputTitle](Excel.Validation.InputTitle.md)
- [Operator](Excel.Validation.Operator.md)
- [Parent](Excel.Validation.Parent.md)
- [ShowError](Excel.Validation.ShowError.md)
- [ShowInput](Excel.Validation.ShowInput.md)
- [Type](Excel.Validation.Type.md)
- [Value](Excel.Validation.Value.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
