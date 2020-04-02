---
title: OLE Programmatic Identifiers (Outlook)
keywords: vbaol11.chm5247509
f1_keywords:
- vbaol11.chm5247509
ms.prod: outlook
ms.assetid: 4dc61073-a674-b786-418e-60b46c79d0c6
ms.date: 06/08/2019
localization_priority: Normal
---


# OLE Programmatic Identifiers (Outlook)

You can use an OLE programmatic identifier (sometimes called a ProgID) to create an **Automation** object. The following tables list OLE programmatic identifiers for ActiveX controls, Microsoft Office applications, and Microsoft Office Web Components.

 [ActiveX Controls](#OLEActiveXControls)

 [Microsoft Access](#OLEMicrosoftAccess)

 [Microsoft Excel](#OLEMicrosoftExcel)
 [Graph](#OLEMicrosoftGraph)
 [Microsoft Outlook](#OLEMicrosoftOutlook)
 [Microsoft PowerPoint](#OLEMicrosoftPowerPoint)
 [Microsoft Word](#OLEMicrosoftWord)

## ActiveX Controls
<a name="OLEActiveXControls"> </a>

To create the ActiveX controls listed in the following table, use the corresponding OLE programmatic identifier.



|**To create this Microsoft Forms 2.0 control**|**Use this identifier**|
|:-----|:-----|
| **CheckBox**|Forms.CheckBox.1|
| **ComboBox**|Forms.ComboBox.1|
| **CommandButton**|Forms.CommandButton.1|
| **Frame**|Forms.Frame.1|
| **Image**|Forms.Image.1|
| **Label**|Forms.Label.1|
| **ListBox**|Forms.ListBox.1|
| **MultiPage**|Forms.MultiPage.1|
| **OptionButton**|Forms.OptionButton.1|
| **ScrollBar**|Forms.ScrollBar.1|
| **SpinButton**|Forms.SpinButton.1|
| **TabStrip**|Forms.TabStrip.1|
| **TextBox**|Forms.TextBox.1|
| **ToggleButton**|Forms.ToggleButton.1|

## Microsoft Access
<a name="OLEMicrosoftAccess"> </a>

To create the Microsoft Access objects listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Access available on the machine where the macro is running.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
| **Application**|Access.Application|
| **CurrentData**|Access.CodeData, Access.CurrentData|
| **CurrentProject**|Access.CodeProject, Access.CurrentProject|
| **DefaultWebOptions**|Access.DefaultWebOptions|

## Microsoft Excel
<a name="OLEMicrosoftExcel"> </a>

To create the Microsoft Excel objects listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Excel available on the machine where the macro is running.



|**To create this object**|**Use one of these identifiers**|**Comments**|
|:-----|:-----|:-----|
| **Application**|Excel.Application||
| **Workbook**|Excel.AddIn||
| **Workbook**|Excel.Chart|Returns a workbook containing two worksheets; one for the chart and one for its data. The chart worksheet is the active worksheet.|
| **Workbook**|Excel.Sheet|Returns a workbook with one worksheet.|

## Graph
<a name="OLEMicrosoftGraph"> </a>

To create the Graph objects listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Graph available on the machine where the macro is running.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
| **Application**|MSGraph.Application|
| **Chart**|MSGraph.Chart|

## Microsoft Outlook
<a name="OLEMicrosoftOutlook"> </a>

To create the Microsoft Outlook object given in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Outlook available on the machine where the macro is running.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
| **[Application](../../../api/Outlook.Application.md)**|Outlook.Application|

To create the ActiveX controls that are specific for Outlook forms, as listed in the following table, use the corresponding OLE programmatic identifier.



|**To create this Outlook control**|**Use this identifier**|
|:-----|:-----|
| **[OlkBusinessCardControl](../../../api/Outlook.OlkBusinessCardControl.md)**|Outlook.OlkBusinessCardControl|
| **[OlkCategory](../../../api/Outlook.OlkCategory.md)**|Outlook.OlkCategoryStrip|
| **[OlkCheckBox](../../../api/Outlook.OlkCheckBox.md)**|Outlook.OlkCheckBox|
| **[OlkComboBox](../../../api/Outlook.OlkComboBox.md)**|Outlook.OlkComboBox|
| **[OlkCommandButton](../../../api/Outlook.OlkCommandButton.md)**|Outlook.OlkCommandButton|
| **[OlkContactPhoto](../../../api/Outlook.OlkContactPhoto.md)**|Outlook.OlkContactPhoto|
| **[OlkDateControl](../../../api/Outlook.OlkDateControl.md)**|Outlook.OlkDateControl|
| **[OlkFrameHeader](../../../api/Outlook.OlkFrameHeader.md)**|Outlook.OlkFrameHeader|
| **[OlkInfoBar](../../../api/Outlook.OlkInfoBar.md)**|Outlook.OlkInfoBar|
| **[OlkLabel](../../../api/Outlook.OlkLabel.md)**|Outlook.OlkLabel|
| **[OlkListBox](../../../api/Outlook.OlkListBox.md)**|Outlook.OlkListBox|
| **[OlkOptionButton](../../../api/Outlook.OlkOptionButton.md)**|Outlook.OlkOptionButton|
| **[OlkPageControl](../../../api/Outlook.OlkPageControl.md)**|Outlook.OlkPageControl|
| **[OlkSenderPhoto](../../../api/Outlook.OlkSenderPhoto.md)**|Outlook.OlkSenderPhoto|
| **[OlkTextBox](../../../api/Outlook.OlkTextBox.md)**|Outlook.OlkTextBox|
| **[OlkTimeControl](../../../api/Outlook.OlkTimeControl.md)**|Outlook.OlkTimeControl|
| **[OlkTimeZoneControl](../../../api/Outlook.OlkTimeZoneControl.md)**|Outlook.OlkTimeZone|

## Microsoft PowerPoint
<a name="OLEMicrosoftPowerPoint"> </a>

To create the Microsoft PowerPoint object given in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of PowerPoint available on the machine where the macro is running.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
| **Application**|PowerPoint.Application|

## Microsoft Word
<a name="OLEMicrosoftWord"> </a>

To create the Microsoft Word objects listed in the following table, use one of the corresponding OLE programmatic identifiers. If you use an identifier without a version number suffix, you create an object in the most recent version of Word available on the machine where the macro is running.



|**To create this object**|**Use one of these identifiers**|
|:-----|:-----|
| **Application**|Word.Application|
| **Document**|Word.Document, Word.Template|
| **Global**|Word.Global|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]