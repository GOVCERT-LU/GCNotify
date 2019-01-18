# Outlook AddIn: GCNotify

Outlook AddIn that creates an Forward as attachment email. This is to ease the work for analyists and users
so that all the information is send whithout hassels.

The Addin will create a Icon in the following ribbons in Outlook

* Home
* NewMail
* ReadMail

![Alt text](/images/outlook_inbox_mod.png?raw=true "Ribbon")

The user has only to select an email from and when selecting the button it creates an email with the selected one attached.
The Fields TO,Body and Subject will be filled automatically. The user sees the newly created email and has to send it manually. This
is so that one can add comments and guarantee a certain transparency.

The code is adapted to our environment and therefore has set the destination email to soc@govcert.etat.lu and there are several GOVCERT.LU references visible to
the user.

The VSTO is tested from Outlook 2013 upwards, but might work on 2010.

# Requirements
Visual Studio 2019 to compile the code

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

# Publication
The solution can be complied via Visual Studio's OneClick Solution, and then distributed.

# LICENSE

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
