---
title: Print, Print Setup dialog boxes
keywords: vbui6.chm2099656
f1_keywords:
- vbui6.chm2099656
ms.prod: office
ms.assetid: 8a81e8c1-21a7-960a-c319-ff4a6ad0c4b0
ms.date: 11/24/2018
localization_priority: Normal
---


# Print, Print Setup dialog boxes

## Print dialog box options

Use the **Print** dialog box to print forms and code to the printer specified in the Control Panel.

|Option|Description|
|:-----|:----------|
|**Printer** |Identifies the printer to which you are printing.|
|**Range**|Determines the range you print:<br/>- **Selection**: Prints the currently selected code.<br/>- **Current Module**: Prints the forms and/or code for the currently selected module.<br/>- **Current Project**: Prints the forms and/or code for the entire project.|
|**Print What**|Determines what you print. You can select as many options as you like, depending on what you selected as the Range.<br/>- **Form Image**: Prints the form images.<br/>- **Code**: Prints the code for the selected range.|
|**Print Quality**|Determines whether you print high, medium, low, or draft output quality.|
|**Print to File**|If selected, print is sent to the file specified in the **Print To File** dialog box. This dialog box appears after you choose OK in the **Print** dialog box.|
|**OK**|Prints your selection.|
|**Cancel**|Closes the dialog box without printing.|
|**Setup**|Displays the standard **Print Setup** dialog box.|


## Print Setup dialog box options

![Print setup dialog box](../../../images/prntset_ZA01201642.gif)

Appears whenever you select **Setup** from the **Print** dialog box.

Use the **Print Setup** dialog box to select the printer, page orientation, and paper size.

|Option|Description|
|:-----|:----------|
|**Printer**|Allows you to specify the printer. If you don't select a printer, Visual Basic prints to the Windows default printer.<br/>- **Name**: Displays a list of available printers.<br/>- **Status**: Displays the status of the printer and whether it is ready to print.<br/>- **Type**: Displays the type of printer.<br/>- **Where**: Displays the location of the printer. If the printer is on a network, displays the path to the server.<br/>- **Comment**: Displays the physical location of the printer and additional information.<br/>- **Properties**: Opens the **Properties** dialog box specific to the printer where you can choose additional options such as paper and the way graphics are printed.|
|**Paper**|Allows you to select the paper size and source (from among those available for the printer). The sizes and sources available depend on the printer you have selected, and they change when you change printers.<br/>- **Size**: Displays a list of the available paper sizes.<br/>- **Source**: Displays the available source of paper for the printer you choose.|
|**Orientation**|Allows you to specify whether the program is to print in Portrait or Landscape orientation.|


## See also

- [Dialog boxes](../dialog-boxes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]