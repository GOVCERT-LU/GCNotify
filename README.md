[![version][version-badge]][CHANGELOG] [![license][license-badge]][LICENSE]

![Alt text](/images/GOVCERT_RGB_for_outlook.png?raw=true "Ribbon" =200px)

# Index
- [GCNotify](#outlook-addin-gcnotify)
  - [Functionalities](#functionalities)
  - [Addin Button Locations](#addin-button-locations)
- [Developpment](#developpment)
  - [Requirements](#requirements)
  - [Customizations](#customizations)
    - [Templates](#templates)
    - [Settings](#settings)
    - [Resources](#resources)
  - [Building](#building)
    - [Compile](#compile)
    - [Publication](#publication)
- [LICENSE](#license)
- [Contribute](#Contribute)
  
# Outlook AddIn: GCNotify

GCNotify is an Outlook AddIn to facilitate the forwarding of suspicious emails to the security team.

It creates a new email with the selected or viewed email as attachment with additional informations (e.g. SMTP Header elements). This is to ease the work for analyists and users:

* For the end user:
  * does not need to create a new email and forward the suspicious email as attachment (Normally this takes several steps).
  * sends it to the right contacts
* For the analyists:
  * does not need to request original mail if it was forwarded without the option "Forward as attachment"
  * adds aditional preprocessed data

This ease of use for the enduser will provide the security team also a greater overview of some threats as the user is more likely to report suspicious emails.

The VSTO is works with Outlook 2013, 2016 and 2019.




## Functionalities
The user has only to select one or more emails from the inbox or opened email and the Addin creates a new email with the selected ones as attachment with a predefined body based on templates.
The destinations and subject of this new email will also be set as specified in the settings of this Addin. The user has only to click on "Send" to send it.

For transpaency purposes the email is not sent automatically. This enables the user to add additional comments to it and also shows what will be send to the security team.



## Addin Button Locations
The Addin will create icons in the following ribbons in Outlook

* Home
* NewMail
* ReadMail
* Send/Receive

![Alt text](/images/outlook_inbox_mod.png?raw=true "Ribbon")


# Developpment
The code is written in Visual Basic

## Requirements
* Visual Studio 2019 enterprise or community edition

## Customizations
The following section describes how GCNotify can be adapted or modified without modifing the source code.
### Settings
In the settings section of the project, the general configurations the following picture shows all the different settings which can be changed and are required for the functionality.

![Alt text](/images/vs_gcnotify_settings.png?raw=true "Visual Studio - Settings")

|**Name**|**Default Value**|**Description**|**Required/Optional**|
|---|---|---|---|
|SOC_MAIL  | soc@govcert.etat.lu  | The main address of the security team. In the email this will be the *TO* field.  |Required|
|SOC_MAIL_CC  |  | The address which should be send a carbon copy to. If left empty ("") it will be ignored. In the email this will be the *CC* field.  |Optional|
|SOC_MAIL_BCC  |  | The address which should be send a blind carbon copy to. If left empty ("") it will be ignored. In the email this will be the *BCC* field. |Optional|
|SUPPORT_MAIL | support@govcert.etat.lu  | The mail address to send errors to. This is when an Exception is thrown. |Required|
|GROUP_LABEL | GOVCERT.LU Tools  | The label of the group in the Ribbon | Required |
|BTN_SUPPERTIP_LABEL | Reports the mail to GOVCERT.LU and requests an analysis  | The label of the supertip, when howering over the button | Required |
|BTN_LABEL | Report Mail  | The label of the button itself | Required |
|INTERESTING_HEADER_FIELDS | Received,Return-Path,X-PMX-Spam,Authentication-Results,Received-SPF,X-Sender,User-Agent,X-Sender,X-Authenticated-Sender,From  | The fields of the header of the email which should be visible in the email. **NOTE**: The values are comma separated. | Required |
|SOC_MAIL_SUBJECT_TAG | [GC-OBT]  | The tag used in the subject | Required |
|SOC_NEW_MAIL_Subject | SOC Request  | The default subject on an empty email | Required |
|SPAM_TAG | SPAM  | The tag used of the email system, when the mail was detected as SPAM. This tag is used to create a confirmation dialog for the user if he really wants to send this email | Required |

**Note:** Required means that it must have a value.

Alternatively they can also be changed in the *app.config* file. This is an XML styled file and a setting is represented as follows

```XML
            <setting name="SPAM_TAG" serializeAs="String">
                <value>SPAM</value>
            </setting>
```

### Templates/Icon
The templates can be found in the resource section of the project or in the **Resources** folder. The text files represent the differemt templates.

|**Filename**|**Description**|**Placeholders**|
|---|---|---|
|EmailDetails.txt|Representation of the extracted information for the forwarded email(s)| **{{EmailCounter}}** - Counter of the attached emails <br/> **{{From}}** - Email Sender <br/>  **{{HeaderDetails}}** - The extracted header informations (Depends on **INTERESTING_HEADER_FIELDS**) <br/>  **{{Subject}}** - Email Subject <br/>  **{{AttachmentCount}}** - Number of attachments in the email  |
|ErrorMail.txt|Email template in case of errors| **{{Version}}** - Version of GCNotify <br/> **{{Message}}** - Message of the exception <br/>  **{{Stacktrace}}** - Stacktrace of the exception |
|NewMailBody.txt|Email template for a new empty email | **{{HostDetails}}** - Details of the host <br/> **{{NetworkDetails}}** - Network details fo the host|
|NewMailBody.txt|Email template for a new empty email | **{{HostDetails}}** - Details of the host <br/> **{{NetworkDetails}}** - Network details fo the host|
|NewResendError.txt| Error Message shown in case of double generation | |
|OverWriteConfirm.txt| Error Message shown if a new email has already been altered an the user hits the button | |
|ResendError.txt| Error Message shown in case of double generation | |
|SPAMDialogText.txt| Error Message shown when a email was already tagged | **{{Email}}** - Email of the sender <br/> **{{Subject}}** - Subject of the email|
|SuspectBody.txt|Email template for a report email | **{{attachments}}** -The place where the email details should be placed (see EmailDetails.txt) <br/> **{{HostDetails}}** - Details of the host <br/> **{{NetworkDetails}}** - Network details fo the host|

**Note:** The Icon can also be changed in a similar faschion.

## Building
MSBuild should be in the PATH variable of Windows, if not it is located here:


> C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin

Go to the folder of the downloaded code and execute:
```
> cd "GOVCERT Outlook Addins"
```
And run:
```
> msbuild "GOVCERT Outlook Addins.vbproj" /t:Publish /p:PublishDir="publish/" /p:Configuration=Release
```
Then the copliled version OneClick Solution should be then found in:

> GOVCERT Outlook Addins\publish

It also can be greated via the Visual Studio's internal publication functionality

![Alt text](/images/vs_gcnotify_publish.png?raw=true "Visual Studio - Pubish")

### Disribution
The compiled solution can be disrtibuted via the OneClick Solution or manually.

#Contribute

Please do contribute! Issues and pull requests are welcome. 


# LICENSE

Copyright (C) 2018, CERT Gouvernemental (GOVCERT.LU)

GC-Notify is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

GC-Notify is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with GC-Notify.  If not, see <https://www.gnu.org/licenses/>.
