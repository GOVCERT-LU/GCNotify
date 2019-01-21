# Outlook Add-in: GC-Notify

The GOVCERT.LU Outlook Add-in allows users to forward messages as attachment to a predefined email address. The aim of the Add-in is to eliminate user errors when submitting suspicious emails and maintaining the format wished by the receiving analysts.

The Add-in will provide an icon in the following Outlook ribbons:

  * Home
  * NewMail
  * ReadMail

![Alt text](/images/outlook_inbox_mod.png?raw=true "Ribbon")

Users who wish to send a suspicious email for analysis will need to select one or more emails that they would like to submit and click the Add-in button in the Outlook ribbon.
The Add-in will open a New Email dialog with its TO, Body and Subject fields pre-filled and the previously selected email(s) attached.
This newly generated email will have to be submitted by the user manually by clicking on the send button, after, if the user wishes to, adding a comment in the body part of the email.

The code of this Add-in has been written in order to meet the requirements of our environment.
The destination address is therefore set to `soc[AT]govcert.etat.lu` and there are several GOVCERT.LU references visible to the user.

The [VSTO](https://en.wikipedia.org/wiki/Visual_Studio_Tools_for_Office) has been tested with Outlook 2013, 2016 and 2019.

# Requirements
 * Visual Studio 2019 to compile the code

# Compile
`MSBuild` should be in the `PATH` variable of Windows, if not it is located here:

```
> C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin
```
co to the folder of the downloaded code and execute:
```
> cd "GOVCERT Outlook Addins"
```
then run:
```
> msbuild "GOVCERT Outlook Addins.vbproj" /t:Publish /p:PublishDir="publish/" /p:Configuration=Release
```

The copiled OneClick Solution version should be then found in:

```
> GOVCERT Outlook Addins\publish
```

# Adaptations

All GOVCERT.LU references can be modified via the solution's properties in
Visual Studio. You can find these in different locations in the properties:

* Application
* Application > Assembly Information
* Resources (Strings, Icons, Files)

However all labels and screen tips, of both the ribbon group and the button, have to be
changed in the following XML files:

* RibbonHome.xml
* RibbonNewMail.xml
* RibbonReadMail.xml

The solution is not signed by default, but we suggest to sign the solution before deployment to automate the installation process, otherwise every user will have to accept to manually install it on their workstation.

Lastly the solution has to be published and distributed through your deployment process.

# Publication
The solution can be complied via Visual Studio's OneClick Solution, and then distributed.

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
