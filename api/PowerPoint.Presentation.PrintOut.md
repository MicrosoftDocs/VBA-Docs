---
title: Presentation.PrintOut method (PowerPoint)
keywords: vbapp10.chm583034
f1_keywords:
- vbapp10.chm583034
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.PrintOut
ms.assetid: 57685390-43c1-4bd4-d2ee-ba34641e34c5
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.PrintOut method (PowerPoint)

Prints the specified presentation.


## Syntax

_expression_.**PrintOut** (_From_, _To_, _PrintToFile_, _Copies_, _Collate_)

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _From_|Optional|**Integer**|The number of the first page to be printed. If this argument is omitted, printing starts at the beginning of the presentation. Specifying the To and From arguments sets the contents of the  **[PrintRanges](PowerPoint.PrintRanges.md)** object and sets the value of the **RangeType** property for the presentation.|
| _To_|Optional|**Integer**|The number of the last page to be printed. If this argument is omitted, printing continues to the end of the presentation. Specifying the To and From arguments sets the contents of the  **[PrintRanges](PowerPoint.PrintRanges.md)** object and sets the value of the **RangeType** property for the presentation.|
| _PrintToFile_|Optional|**String**|The name of the file to print to. If you specify this argument, the file is printed to a file rather than sent to a printer. If this argument is omitted, the file is sent to a printer.|
| _Copies_|Optional|**Integer**|The number of copies to be printed. If this argument is omitted, only one copy is printed. Specifying this argument sets the value of the [NumberOfCopies](PowerPoint.PrintOptions.NumberOfCopies.md)property.|
| _Collate_|Optional|**MsoTriState**|If this argument is omitted, multiple copies are collated. Specifying this argument sets the value of the  **[Collate](PowerPoint.PrintOptions.Collate.md)** property.|

## Remarks

The  _Collate_ parameter value can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|Prints all copies of one page before printing the first copy of the next page.|
|**msoTrue**|Prints a complete copy of the presentation before the first page of the next copy is printed.|

## Example

This example prints two uncollated copies of each slide&mdash;whether visible or hidden&mdash;from slide two to slide five in the active presentation.


```vb
With Application.ActivePresentation

    .PrintOptions.PrintHiddenSlides = True

    .PrintOut From:=2, To:=5, Copies:=2, Collate:=msoFalse

End With


```

This example prints a single copy of all slides in the active presentation to the file Testprnt.prn.




```vb
Application.ActivePresentation.PrintOut PrintToFile:="TestPrnt"
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]