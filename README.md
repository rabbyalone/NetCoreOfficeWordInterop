# NetCoreOfficeWordInterop
Find Replace from doc template, and Export As PDF.

### Pre-Requisite

* MS Office 2010 and up installed on dev machine or deployed server.

### For New project with COM object in .NET Core

The CSPROJ file, If anyone face problem setting up COM Reference on ASP NEt Core application.

```
<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>netcoreapp3.0</TargetFramework>
  </PropertyGroup>

  <ItemGroup>
    <Content Remove="wwwroot\lib\Interop.Microsoft.Office.Interop.Word.dll" />
  </ItemGroup>

  <ItemGroup>
    <COMReference Include="Microsoft.Office.Word.dll">
      <Guid>00020905-0000-0000-c000-000000000046</Guid>
      <VersionMajor>8</VersionMajor>
      <VersionMinor>6</VersionMinor>
      <WrapperTool>primary</WrapperTool>
      <Lcid>0</Lcid>
      <Isolated>false</Isolated>
      <Private>false</Private>
      <EmbedInteropTypes>true</EmbedInteropTypes>
      <Aliases>global</Aliases>
    </COMReference>
  </ItemGroup>

  <ItemGroup>
    <Resource Include="wwwroot\lib\Interop.Microsoft.Office.Interop.Word.dll">
      <CopyToOutputDirectory>Always</CopyToOutputDirectory>
    </Resource>
  </ItemGroup>
  
```


