<img src="/images/GOVCERT_RGB_for_outlook.png?raw=true"  width="200" height="200">

# Index
- [GCNotify](#outlook-add-in-gcnotify)
  - [Functionalities](#functionalities)
  - [Features](#features)
  - [Add-in Button Locations](#add-in-button-locations)
  - [Example Email](#example-email)
- [Development](#development)
  - [Requirements](#requirements)
  - [Customisations](#customisations)
    - [Templates/Icon](#templatesicon)
    - [Settings](#settings)
  - [Building](#building)
    - [Distribution](#distribution)
- [Contribute](#contribute)
- [LICENSE](#license)
  
# Outlook Add-in: GCNotify

GCNotify is an Outlook Add-in to facilitate the forwarding of suspicious emails to an IT-security team.

It creates a new email with the selected or viewed email as attachment with additional informations (e.g. SMTP Header elements). This is to ease the work of security analysts and users:

* For the end user:
  * does not need to forward the suspicious email as attachment as a new email (this would take multiple steps).
  * sends it to the right IT-security team addresses
* For the analysts:
  * no need to request the original email to be forwarded as attachment
  * additional preprocessed data added

This ease of use for the end user will provide the IT-security team to have a greater overview of threats as user are more likely to report suspicious emails.

The VSTO is works with Outlook 2013, 2016 and 2019.




## Functionalities
The user has to select one or more emails from their inbox or an opened email.
The Add-in creates a new email with the selected emails as attachment and adds a predefined body based on templates.
The destination and subject of this new email will also be pre-filled as defined in the settings of the Add-in.
The only action a user has to trigger is hit "Send".

For transparency purposes the email is not sent without the users consent.
This also allows the user to add additional comments and also displays what will be send to the IT-security team.

## Features
* Easy to use
* Sends one or multiple emails as attachment
* Customisable

## Add-in Button Locations
The Add-in will add icons in the following ribbons in Outlook

* Home
* NewMail
* ReadMail
* Send/Receive

![Alt text](/images/outlook_inbox_mod.png?raw=true "Ribbon")

## Example Email
<img src="/images/Outlook_mail.png?raw=true"  width="250" height="250">

# Development
The code is written in Visual Basic

## Requirements
* Visual Studio 2019 enterprise or community edition

## Customisations
The following section describes how GCNotify can be adjusted to your needs without modifying the source code.

### Settings
The settings section in Visual Studio allows you to adjust GCNotify.
This section describes the different settings and which ones are required for the plug-in to work.

![Alt text](/images/vs_gcnotify_settings.png?raw=true "Visual Studio - Settings")

|**Name**|**Default Value**|**Description**|**Required/Optional**|
|--------|-----------------|---------------|---------------------|
|SOC_MAIL                   | soc@govcert.etat.lu       | The main email address of the IT-security team. In the generated email this will be the `TO` field.  |Required|
|SOC_MAIL_CC                |                           | Email address which should receive a carbon copy. If left empty ('') it will be ignored. In the generated email this will be the `CC` field.  |Optional|
|SOC_MAIL_BCC               |                           | Email address which should receive a blind carbon copy. If left empty ('') it will be ignored. In the generated email this will be the `BCC` field. |Optional|
|SUPPORT_MAIL               | support@govcert.etat.lu   | The email address to send errors to. This destination is used when an Exception is thrown. |Required|
|GROUP_LABEL                | GOVCERT.LU Tools          | The label of the ribbon group | Required |
|BTN_SUPPERTIP_LABEL        | Reports the mail to GOVCERT.LU and requests an analysis | The label of the supertip, when hovering over the button | Required |
|BTN_LABEL                  | Report Mail               | The label of the button itself | Required |
|INTERESTING_HEADER_FIELDS  | Received,Return-Path,X-PMX-Spam,Authentication-Results,Received-SPF,X-Sender,User-Agent,X-Sender,X-Authenticated-Sender,From  | The header fields of the email which should be visible in the email. **NOTE**: The values are comma separated. | Required |
|SOC_MAIL_SUBJECT_TAG       | [GC-OBT]                  | The tag used in the subject | Required |
|SOC_NEW_MAIL_Subject       | SOC Request               | The default subject of an empty email | Required |
|SPAM_TAG                   | SPAM                      | The tag used of the email system, when the mail was detected as SPAM. This tag is used to open a confirmation dialog in order to make sure the user really wants to send this email | Required |

**Note:** Required means that the setting must not be empty.

Alternatively they can also be changed in the *app.config* file. This is an XML file where settings are represented as follows:

```XML
            <setting name="SPAM_TAG" serializeAs="String">
                <value>SPAM</value>
            </setting>
```

### Templates/Icon
The templates can be found in the resource section of the project or in the **Resources** folder. The text files represent the different templates.

|**Filename**|**Description**|**Placeholders**|
|------------|---------------|----------------|
|EmailDetails.txt     | Representation of the extracted information of forwarded email(s)| **{{EmailCounter}}** - Index of attached emails <br/> **{{From}}** - Email sender <br/>  **{{HeaderDetails}}** - The extracted header information (Depends on **INTERESTING_HEADER_FIELDS**) <br/>  **{{Subject}}** - Email subject <br/>  **{{AttachmentCount}}** - Amount of attachments in the email  |
|ErrorMail.txt        | Email body template in case of an exception| **{{Version}}** - Version of GCNotify <br/> **{{Message}}** - Exception message <br/>  **{{Stacktrace}}** - Exception stack trace |
|NewMailBody.txt      | Email body template for a new empty email | **{{HostDetails}}** - Details of the host <br/> **{{NetworkDetails}}** - Network details fo the host|
|OverWriteConfirm.txt | Message displayed if a user has opened a new email window, filled in some content and then hit the button. In order not to overwrite the information a user has already entered, the user is asked whether this information shall be overwritten or not. | |
|ResendError.txt      | Message displayed in case a user hits the GCNotify button in the composing window of the reporting email | |
|SPAMDialogText.txt   | Message displayed when an email tagged as spam is within the selection of the emails to be forwarded | **{{Email}}** - Email of the sender <br/> **{{Subject}}** - Subject of the email|
|SuspectBody.txt      | Email body template for a report email | **{{attachments}}** - The place where the email details should be placed within the body (see EmailDetails.txt) <br/> **{{HostDetails}}** - Details of the host <br/> **{{NetworkDetails}}** - Network details of the host|

**Note:** The Icon can also be changed in a similar fashion.

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
Then the compiled OneClick Solution should be now be located in:

> GOVCERT Outlook Addins\publish

It can also be generated via Visual Studio's internal publication functionality

![Alt text](/images/vs_gcnotify_publish.png?raw=true "Visual Studio - Publish")


### Distribution
The project should be signed; this can be configured in the properties / signing tab. If you want the plugin to outlive your certificate's validity period, you should consider setting up timestamping ("Timestamp server URL" field).

The compiled solution can be distributed via the OneClick Solution or manually.



# Contribute

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
