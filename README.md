[![version][version-badge]][CHANGELOG] [![license][license-badge]][LICENSE]

- [GCNotify](#outlook-addin-gcnotify)
  - [Functionalities](#functionalities)
  - [Addin Button Locations](#Addin-Button-Locations)
- [Developpment](#Developpment)
  - [Requirements](#Requirements)
  - [Customizations](#Customizations)
  - [Building](#Building)
  - [Publication](#Publication)
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
Visual Studio 2019 to compile the code
## Customizations
 - app.config
 - Settings in VS
 - 
### Templates
### Settings
### Resources

## Building
# Compile
MSBuild should be in the PATH variable of Windows, if not it is located here:

```
> C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin
```
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

# Adaptations

The GOVCERT.LU references can be changed via the solution's Properties in
Visual Studio. The following list shows the different locations in the
properties:

* Application
* Application > Assembly Information
* Resources (Strings, Icons, Files)

However the labels/screen tips of the ribbon group and button have to be
changed in the following XML files:

* RibbonHome.xml
* RibbonNewMail.xml
* RibbonReadMail.xml

The solution, by default, is not signed but it is suggested to add one known to the destined workstations to automate the installation process, else
the user has to click on install manually to accept it. 


Then the solution has to be published and distributed.

## Publication
The solution can be complied via Visual Studio's OneClick Solution, and then distributed.


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
