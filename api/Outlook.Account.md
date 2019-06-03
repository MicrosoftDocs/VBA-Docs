---
title: Account object (Outlook)
keywords: vbaol11.chm3153
f1_keywords:
- vbaol11.chm3153
ms.prod: outlook
api_name:
- Outlook.Account
ms.assetid: f624438c-4e45-2822-18b6-bfe8074a33c0
ms.date: 06/08/2017
localization_priority: Normal
---


# Account object (Outlook)

The  **Account** object represents an account that is defined for the current profile.


## Remarks

The purpose of the [Accounts](Outlook.Accounts.md) collection object and the **Account** object is to provide the capacity to enumerate **Account** objects in a given profile, to identify the type of **Account**, and to use a specific **Account** object to send mail.


> [!NOTE] 
> Helmut Obertanner provided the following code samples. Helmut is a [Microsoft Most Valuable Professional](https://mvp.microsoft.com/) with expertise in Microsoft Office development tools in Microsoft Visual Studio and Microsoft Office Outlook.


## Example

The following managed code samples are written in C# and Visual Basic. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code samples in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The following code samples show the  `DisplayAccountInformation` method of the `Sample` class, implemented as part of an Outlook add-in project. Each project adds a reference to the Outlook PIA, which is based on the **Microsoft.Office.Interop.Outlook** namespace. The `DisplayAccountInformation` method takes as an input argument a trusted Outlook[Application](Outlook.Application.md) object, and uses the **Account** object to display the details of each account that is available for the current Outlook profile.




```cs
using System; 
using System.Text; 
using Outlook = Microsoft.Office.Interop.Outlook; 
 
namespace OutlookAddIn1 
{ 
 class Sample 
 { 
 public static void DisplayAccountInformation(Outlook.Application application) 
 { 
 
 // The Namespace Object (Session) has a collection of accounts. 
 Outlook.Accounts accounts = application.Session.Accounts; 
 
 // Concatenate a message with information about all accounts. 
 StringBuilder builder = new StringBuilder(); 
 
 // Loop over all accounts and print detail account information. 
 // All properties of the Account object are read-only. 
 foreach (Outlook.Account account in accounts) 
 { 
 
 // The DisplayName property represents the friendly name of the account. 
 builder.AppendFormat("DisplayName: {0}\n", account.DisplayName); 
 
 // The UserName property provides an account-based context to determine identity. 
 builder.AppendFormat("UserName: {0}\n", account.UserName); 
 
 // The SmtpAddress property provides the SMTP address for the account. 
 builder.AppendFormat("SmtpAddress: {0}\n", account.SmtpAddress); 
 
 // The AccountType property indicates the type of the account. 
 builder.Append("AccountType: "); 
 switch (account.AccountType) 
 { 
 
 case Outlook.OlAccountType.olExchange: 
 builder.AppendLine("Exchange"); 
 break; 
 
 case Outlook.OlAccountType.olHttp: 
 builder.AppendLine("Http"); 
 break; 
 
 case Outlook.OlAccountType.olImap: 
 builder.AppendLine("Imap"); 
 break; 
 
 case Outlook.OlAccountType.olOtherAccount: 
 builder.AppendLine("Other"); 
 break; 
 
 case Outlook.OlAccountType.olPop3: 
 builder.AppendLine("Pop3"); 
 break; 
 } 
 
 builder.AppendLine(); 
 } 
 
 // Display the account information. 
 System.Windows.Forms.MessageBox.Show(builder.ToString()); 
 } 
 } 
}
```




```vb
Imports Outlook = Microsoft.Office.Interop.Outlook 
 
Namespace OutlookAddIn2 
 Class Sample 
 Shared Sub DisplayAccountInformation(ByVal application As Outlook.Application) 
 
 ' The Namespace Object (Session) has a collection of accounts. 
 Dim accounts As Outlook.Accounts = application.Session.Accounts 
 
 ' Concatenate a message with information about all accounts. 
 Dim builder As StringBuilder = New StringBuilder() 
 
 ' Loop over all accounts and print detail account information. 
 ' All properties of the Account object are read-only. 
 Dim account As Outlook.Account 
 For Each account In accounts 
 
 ' The DisplayName property represents the friendly name of the account. 
 builder.AppendFormat("DisplayName: {0}" & vbNewLine, account.DisplayName) 
 
 ' The UserName property provides an account-based context to determine identity. 
 builder.AppendFormat("UserName: {0}" & vbNewLine, account.UserName) 
 
 ' The SmtpAddress property provides the SMTP address for the account. 
 builder.AppendFormat("SmtpAddress: {0}" & vbNewLine, account.SmtpAddress) 
 
 ' The AccountType property indicates the type of the account. 
 builder.Append("AccountType: ") 
 Select Case (account.AccountType) 
 
 Case Outlook.OlAccountType.olExchange 
 builder.AppendLine("Exchange") 
 
 
 Case Outlook.OlAccountType.olHttp 
 builder.AppendLine("Http") 
 
 
 Case Outlook.OlAccountType.olImap 
 builder.AppendLine("Imap") 
 
 
 Case Outlook.OlAccountType.olOtherAccount 
 builder.AppendLine("Other") 
 
 
 Case Outlook.OlAccountType.olPop3 
 builder.AppendLine("Pop3") 
 
 
 End Select 
 
 builder.AppendLine() 
 Next 
 
 
 ' Display the account information. 
 Windows.Forms.MessageBox.Show(builder.ToString()) 
 End Sub 
 
 
 End Class 
End Namespace
```


## Methods



|Name|
|:-----|
|[GetAddressEntryFromID](Outlook.Account.GetAddressEntryFromID.md)|
|[GetRecipientFromID](Outlook.Account.GetRecipientFromID.md)|

## Properties



|Name|
|:-----|
|[AccountType](Outlook.Account.AccountType.md)|
|[Application](Outlook.Account.Application.md)|
|[AutoDiscoverConnectionMode](Outlook.Account.AutoDiscoverConnectionMode.md)|
|[AutoDiscoverXml](Outlook.Account.AutoDiscoverXml.md)|
|[Class](Outlook.Account.Class.md)|
|[CurrentUser](Outlook.Account.CurrentUser.md)|
|[DeliveryStore](Outlook.Account.DeliveryStore.md)|
|[DisplayName](Outlook.Account.DisplayName.md)|
|[ExchangeConnectionMode](Outlook.Account.ExchangeConnectionMode.md)|
|[ExchangeMailboxServerName](Outlook.Account.ExchangeMailboxServerName.md)|
|[ExchangeMailboxServerVersion](Outlook.Account.ExchangeMailboxServerVersion.md)|
|[Parent](Outlook.Account.Parent.md)|
|[Session](Outlook.Account.Session.md)|
|[SmtpAddress](Outlook.Account.SmtpAddress.md)|
|[UserName](Outlook.Account.UserName.md)|

## See also


[Account Object Members](overview/Outlook.md)
[Send an email given the SMTP address of an account](../outlook/How-to/Items-Folders-and-Stores/send-an-e-mail-given-the-smtp-address-of-an-account-outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
