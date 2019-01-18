Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security

' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes

<Assembly: AssemblyTitle("GOVCERT.LU Outlook Addins")>
<Assembly: AssemblyDescription("GOVCERT.LU Support Tools")>
<Assembly: AssemblyCompany("CERT Gouvernemental Luxembourg")>
<Assembly: AssemblyProduct("GOVCERT.LU Outlook Addins")>
<Assembly: AssemblyCopyright("Copyright (C) 2018, CERT Gouvernemental (GOVCERT.LU)")>
<Assembly: AssemblyTrademark("")>

' Setting ComVisible to false makes the types in this assembly not visible 
' to COM components.  If you need to access a type in this assembly from 
' COM, set the ComVisible attribute to true on that type.
<Assembly: ComVisible(True)>

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("8c8908d6-fe89-4f6e-b10e-cee306cbb067")>

' Version information for an assembly consists of the following four values:
'
'      Major Version
'      Minor Version 
'      Build Number
'      Revision
'
' You can specify all the values or you can default the Build and Revision Numbers 
' by using the '*' as shown below:
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.4.0.0")>
<Assembly: AssemblyFileVersion("1.4.0.0")>

Friend Module DesignTimeConstants
    Public Const RibbonTypeSerializer As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Public Const RibbonBaseTypeSerializer As String = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design"
    Public Const RibbonDesigner As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
End Module
