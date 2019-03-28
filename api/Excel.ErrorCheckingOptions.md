---
title: ErrorCheckingOptions object (Excel)
keywords: vbaxl10.chm697072
f1_keywords:
- vbaxl10.chm697072
ms.prod: excel
api_name:
- Excel.ErrorCheckingOptions
ms.assetid: f62d3b08-a08f-d028-8e33-4bfd8799dc44
ms.date: 03/29/2019
localization_priority: Normal
---


# ErrorCheckingOptions object (Excel)

Represents the error-checking options for an application.


## Remarks

Use the **[ErrorCheckingOptions](Excel.Application.ErrorCheckingOptions.md)** property of the **Application** object to return an **ErrorCheckingOptions** object.

Reference the **[Item](Excel.Errors.Item.md)** property of the **Errors** object to view a list of index values associated with error-checking options.

After an **ErrorCheckingOptions** object is returned, you can use the following properties, which are members of the **ErrorCheckingOptions** object, to set or return error checking options.

- **BackgroundChecking**   
- **EmptyCellReferences**   
- **EvaluateToError**   
- **InconsistentFormula**   
- **IndicatorColorIndex**  
- **NumberAsText**   
- **OmittedCells**   
- **TextDate**  
- **UnlockedFormulaCells**
    

## Example

The following example uses the **TextDate** property to enable error checking for two-digit-year text dates, and then notifies the user.

```vb
Sub CheckTextDates() 
 
 Dim rngFormula As Range 
 Set rngFormula = Application.Range("A1") 
 
 Range("A1").Formula = "'April 23, 00" 
 Application.ErrorCheckingOptions.TextDate = True 
 
 ' Perform check to see if 2 digit year TextDate check is on. 
 If rngFormula.Errors.Item(xlTextDate).Value = True Then 
 MsgBox "The text date error checking feature is enabled." 
 Else 
 MsgBox "The text date error checking feature is not on." 
 End If 
 
End Sub
```

## Properties

- [Application](Excel.ErrorCheckingOptions.Application.md)
- [BackgroundChecking](Excel.ErrorCheckingOptions.BackgroundChecking.md)
- [Creator](Excel.ErrorCheckingOptions.Creator.md)
- [EmptyCellReferences](Excel.ErrorCheckingOptions.EmptyCellReferences.md)
- [EvaluateToError](Excel.ErrorCheckingOptions.EvaluateToError.md)
- [InconsistentFormula](Excel.ErrorCheckingOptions.InconsistentFormula.md)
- [InconsistentTableFormula](Excel.ErrorCheckingOptions.InconsistentTableFormula.md)
- [IndicatorColorIndex](Excel.ErrorCheckingOptions.IndicatorColorIndex.md)
- [ListDataValidation](Excel.ErrorCheckingOptions.ListDataValidation.md)
- [NumberAsText](Excel.ErrorCheckingOptions.NumberAsText.md)
- [OmittedCells](Excel.ErrorCheckingOptions.OmittedCells.md)
- [Parent](Excel.ErrorCheckingOptions.Parent.md)
- [TextDate](Excel.ErrorCheckingOptions.TextDate.md)
- [UnlockedFormulaCells](Excel.ErrorCheckingOptions.UnlockedFormulaCells.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]