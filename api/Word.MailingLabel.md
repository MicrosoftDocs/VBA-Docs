---
title: MailingLabel object (Word)
keywords: vbawd10.chm2327
f1_keywords:
- vbawd10.chm2327
ms.prod: word
api_name:
- Word.MailingLabel
ms.assetid: 9dd073b7-5d53-0f1e-f19a-9abf6427b3f2
ms.date: 06/08/2017
localization_priority: Normal
---


# MailingLabel object (Word)

Represents a mailing label.


## Remarks

Use the **MailingLabel** property to return the **MailingLabel** object. The following example sets default mailing label options.


```vb
With Application.MailingLabel 
 .DefaultLaserTray = wdPrinterLowerBin 
 .DefaultPrintBarCode = True 
End With
```

Use the **PrintOut** method to print a mailing label listed in the **Product Number** box in the **Label Options** dialog box. The following example prints a page of Avery 5162 standard address labels using the specified address.




```vb
addr = "Katie Jordan" & vbCr & "123 Skye St." _ 
 & vbCr & "OurTown, WA 98107" 
Application.MailingLabel.PrintOut Name:="5162", Address:=addr
```

Use the **CustomLabels** property to format or print a custom mailing label. The following example sets the number of labels across and down for the custom label named "MyLabel."




```vb
With Application.MailingLabel.CustomLabels("MyLabel") 
 .NumberAcross = 2 
 .NumberDown = 5 
End With
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]