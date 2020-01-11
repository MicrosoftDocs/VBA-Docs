---
title: Obtain the Email Address of a Recipient
ms.prod: outlook
ms.assetid: b645c227-a7d2-2861-3bf7-4190a19abe81
ms.date: 06/08/2017
localization_priority: Normal
---


# Obtain the Email Address of a Recipient

This topic shows how to obtain the SMTP address for each recipient in a  **[Recipients](../../../api/Outlook.Recipients.md)** collection.

The method in the code sample,  `GetSMTPAddressForRecipients`, takes a  **[MailItem](../../../api/Outlook.MailItem.md)** as an input argument and then displays the SMTP address of each recipient for that mail item. The method first retrieves the **Recipients** collection that represents the set of recipients specified for the mail item. For each **[Recipient](../../../api/Outlook.Recipient.md)** in that **Recipients** collection, the method then obtains the **[PropertyAccessor](../../../api/Outlook.PropertyAccessor.md)** object that corresponds to that **Recipient** object, and uses the **PropertyAccessor** to get the value of the MAPI property `https://schemas.microsoft.com/mapi/proptag/0x39FE001E`, that maps to the SMTP address of the recipient.

This topic contains two code samples. The following code sample is written in Microsoft Visual Basic for Applications (VBA). 




```vb
Sub GetSMTPAddressForRecipients(mail As Outlook.MailItem) 
    Dim recips As Outlook.Recipients 
    Dim recip As Outlook.Recipient 
    Dim pa As Outlook.PropertyAccessor 
    Const PR_SMTP_ADDRESS As String = _ 
        "https://schemas.microsoft.com/mapi/proptag/0x39FE001E" 
    Set recips = mail.Recipients 
    For Each recip In recips 
        Set pa = recip.PropertyAccessor 
        Debug.Print recip.name & " SMTP=" _ 
           & pa.GetProperty(PR_SMTP_ADDRESS) 
    Next 
End Sub
```

The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.



```cs
private void GetSMTPAddressForRecipients(Outlook.MailItem mail) 
{ 
    const string PR_SMTP_ADDRESS = 
        "https://schemas.microsoft.com/mapi/proptag/0x39FE001E"; 
    Outlook.Recipients recips = mail.Recipients; 
    foreach (Outlook.Recipient recip in recips) 
    { 
        Outlook.PropertyAccessor pa = recip.PropertyAccessor; 
        string smtpAddress = 
            pa.GetProperty(PR_SMTP_ADDRESS).ToString(); 
        Debug.WriteLine(recip.Name + " SMTP=" + smtpAddress); 
    } 
} 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
