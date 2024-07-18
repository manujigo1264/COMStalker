# COMStalker

COMStalker is a tool designed to enumerate and analyze COM (Component Object Model) servers registered on a Windows system. It provides detailed information about the COM servers, including their CLSID, server path, type, and methods exposed by .NET assemblies.

## Features

- Enumerates InprocServer32 and LocalServer32 COM servers.
- Displays CLSID, server path, and server type.
- Identifies and lists methods in .NET assemblies.

## Usage

```sh
COMStalker.exe <-inproc|-localserver>

# Example Output

COMStalker by 0xTron

COM Servers Information
=======================
CLSID: {000209FF-0000-0000-C000-000000000046}
Path: C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE
Type: LocalServer32
-----------------------
.NET Assembly: file:///C:/Program Files/SomeApp/SomeApp.dll
  Method: VulnerableMethod
=======================
