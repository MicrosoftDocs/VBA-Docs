---
title: ErrorCheckingOptions Object (Excel)
keywords: vbaxl10.chm697072
f1_keywords:
- vbaxl10.chm697072
ms.prod: excel
api_name:
- Excel.ErrorCheckingOptions
ms.assetid: f62d3b08-a08f-d028-8e33-4bfd8799dc44
ms.date: 06/08/2017
---


# ErrorCheckingOptions Object (Excel)

Represents the error-checking options for an application.


## Remarks

Use the  **[ErrorCheckingOptions](Excel.Application.ErrorCheckingOptions.md)** property of the **[Application](Excel.Application(objec).md)** object to return an **ErrorCheckingOptions** object.

Reference the  **[Item](Excel.Errors.Item.md)** property of the **[Errors](Excel.Errors.md)** object to view a list of index values associated with error-checking options.

Once an  **ErrorCheckingOptions** object is returned, you can use the following properties, which are members of the **ErrorCheckingOptions** object, to set or return error checking options.


-  **[BackgroundChecking](Excel.ErrorCheckingOptions.BackgroundChecking.md)**
    
-  **[EmptyCellReferences](Excel.ErrorCheckingOptions.EmptyCellReferences.md)**
    
-  **[EvaluateToError](Excel.ErrorCheckingOptions.EvaluateToError.md)**
    
-  **[InconsistentFormula](Excel.ErrorCheckingOptions.InconsistentFormula.md)**
    
-  **[IndicatorColorIndex](Excel.ErrorCheckingOptions.IndicatorColorIndex.md)**
    
-  **[NumberAsText](Excel.ErrorCheckingOptions.NumberAsText.md)**
    
-  **[OmittedCells](Excel.ErrorCheckingOptions.OmittedCells.md)**
    
-  **[TextDate](Excel.ErrorCheckingOptions.TextDate.md)**
    
-  **[UnlockedFormulaCells](Excel.ErrorCheckingOptions.UnlockedFormulaCells.md)**
    

## Example

The following example uses the  **TextDate** property to enable error checking for two-digit-year text dates and notifies the user.


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


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

