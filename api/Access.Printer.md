---
title: Printer object (Access)
keywords: vbaac10.chm12880
f1_keywords:
- vbaac10.chm12880
ms.prod: access
api_name:
- Access.Printer
ms.assetid: fba3eb15-db93-943a-421c-291761e7fa2b
ms.date: 03/21/2019
localization_priority: Normal
---


# Printer object (Access)

A **Printer** object corresponds to a printer available on your system.


## Remarks

A **Printer** object is a member of the **[Printers](Access.Printers.md)** collection.

To return a reference to a particular **Printer** object in the **Printers** collection, use any of the following syntax forms.

|Syntax|Description|
|:-----|:-----|
|**Printers**!_devicename_|The _devicename_ argument is the name of the **Printer** object as returned by the **DeviceName** property.|
|**Printers**("_devicename_")|The _devicename_ argument is the name of the **Printer** object as returned by the **DeviceName** property.|
|**Printers**(_index_)|The _index_ argument is the numeric position of the object within the collection. The valid range is from 0 to `Printers.Count-1`.|

You can use the properties of the **Printer** object to set the printing characteristics for any of the printers available on your system.

Use the **ColorMode**, **Copies**, **Duplex**, **Orientation**, **PaperBin**, **PaperSize**, and **PrintQuality** properties to specify print settings for a particular printer.

Use the **LeftMargin**, **RightMargin**, **TopMargin**, **BottomMargin**, **ColumnSpacing**, **RowSpacing**, **DataOnly**, **DefaultSize**, **ItemLayout**, **ItemsAcross**, **ItemSizeHeight**, and **ItemSizeWidth** properties to specify how Microsoft Access should format the appearance of data on printed pages.

Use the **DeviceName**, **DriverName**, and **Port** properties to return system information about a particular printer.


## Example

The following example displays system information about the first printer in the **Printers** collection.

```vb
Dim prtFirst As Printer 
 
Set prtFirst = Application.Printers(0) 
 
With prtFirst 
 MsgBox "Device name: " & .DeviceName & vbCr _ 
 & "Driver name: " & .DriverName & vbCr _ 
 & "Port: " & .Port 
End With
```


## Properties

- [BottomMargin](Access.Printer.BottomMargin.md)
- [ColorMode](Access.Printer.ColorMode.md)
- [ColumnSpacing](Access.Printer.ColumnSpacing.md)
- [Copies](Access.Printer.Copies.md)
- [DataOnly](Access.Printer.DataOnly.md)
- [DefaultSize](Access.Printer.DefaultSize.md)
- [DeviceName](Access.Printer.DeviceName.md)
- [DriverName](Access.Printer.DriverName.md)
- [Duplex](Access.Printer.Duplex.md)
- [ItemLayout](Access.Printer.ItemLayout.md)
- [ItemsAcross](Access.Printer.ItemsAcross.md)
- [ItemSizeHeight](Access.Printer.ItemSizeHeight.md)
- [ItemSizeWidth](Access.Printer.ItemSizeWidth.md)
- [LeftMargin](Access.Printer.LeftMargin.md)
- [Orientation](Access.Printer.Orientation.md)
- [PaperBin](Access.Printer.PaperBin.md)
- [PaperSize](Access.Printer.PaperSize.md)
- [Port](Access.Printer.Port.md)
- [PrintQuality](Access.Printer.PrintQuality.md)
- [RightMargin](Access.Printer.RightMargin.md)
- [RowSpacing](Access.Printer.RowSpacing.md)
- [TopMargin](Access.Printer.TopMargin.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
