[Setup]
AppName=Enhanced Features
AppId=EnhancedFeatures
AppVerName=Enhanced Features 0.9.3.0
AppCopyright=Copyright © Doena Soft. 2017
AppPublisher=Doena Soft.
AppPublisherURL=http://doena-journal.net/en/dvd-profiler-tools/
DefaultDirName={pf32}\Doena Soft.\Enhanced Features
; DefaultGroupName=Doena Soft.
DirExistsWarning=No
SourceDir=..\EnhancedFeatures\bin\x86\EnhancedFeatures
Compression=zip/9
AppMutex=InvelosDVDPro
OutputBaseFilename=EnhancedFeaturesSetup
OutputDir=..\..\..\..\EnhancedFeaturesSetup\Setup\EnhancedFeatures
MinVersion=0,5.1
PrivilegesRequired=admin
WizardImageFile=compiler:wizmodernimage-is.bmp
WizardSmallImageFile=compiler:wizmodernsmallimage-is.bmp
DisableReadyPage=yes
ShowLanguageDialog=no
VersionInfoCompany=Doena Soft.
VersionInfoCopyright=2017
VersionInfoDescription=Enhanced Features Setup
VersionInfoVersion=0.9.3.0
UninstallDisplayIcon={app}\djdsoft.ico

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
WinVersionTooLowError=This program requires Windows XP or above to be installed.%n%nWindows 9x, NT and 2000 are not supported.

[Types]
Name: "full"; Description: "Full installation"

[Files]
Source: "djdsoft.ico"; DestDir: "{app}"; Flags: ignoreversion

Source: "EnhancedFeatures.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "EnhancedFeatures.pdb"; DestDir: "{app}"; Flags: ignoreversion

Source: "EnhancedFeatures.xsd"; DestDir: "{app}"; Flags: ignoreversion

Source: "DVDProfilerHelper.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "DVDProfilerHelper.pdb"; DestDir: "{app}"; Flags: ignoreversion

Source: "EnhancedFeaturesLibrary.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "EnhancedFeaturesLibrary.pdb"; DestDir: "{app}"; Flags: ignoreversion

Source: "Microsoft.WindowsAPICodePack.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "Microsoft.WindowsAPICodePack.Shell.dll"; DestDir: "{app}"; Flags: ignoreversion

Source: "de\EnhancedFeatures.resources.dll"; DestDir: "{app}\de"; Flags: ignoreversion
Source: "de\DVDProfilerHelper.resources.dll"; DestDir: "{app}\de"; Flags: ignoreversion

Source: "fr\EnhancedFeatures.resources.dll"; DestDir: "{app}\fr"; Flags: ignoreversion

; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Run]
Filename: "{win}\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe"; Parameters: "/codebase ""{app}\EnhancedFeatures.dll"""; Flags: runhidden

;[UninstallDelete]

[UninstallRun]
Filename: "{win}\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe"; Parameters: "/u ""{app}\EnhancedFeatures.dll"""; Flags: runhidden

[Registry]
; Register - Cleanup ahead of time in case the user didn't uninstall the previous version.
Root: HKCR; Subkey: "CLSID\{{58E2706B-250B-4ED4-BD69-82DEA5A1BBD6}"; Flags: dontcreatekey deletekey
Root: HKCR; Subkey: "DoenaSoft.DVDProfiler.EnhancedFeatures.Plugin"; Flags: dontcreatekey deletekey
Root: HKCU; Subkey: "Software\Invelos Software\DVD Profiler\Plugins\Identified"; ValueType: none; ValueName: "{{58E2706B-250B-4ED4-BD69-82DEA5A1BBD6}"; ValueData: "0"; Flags: deletevalue
Root: HKCU; Subkey: "Software\Invelos Software\DVD Profiler_beta\Plugins\Identified"; ValueType: none; ValueName: "{{58E2706B-250B-4ED4-BD69-82DEA5A1BBD6}"; ValueData: "0"; Flags: deletevalue
Root: HKLM; Subkey: "Software\Classes\CLSID\{{58E2706B-250B-4ED4-BD69-82DEA5A1BBD6}"; Flags: dontcreatekey deletekey
Root: HKLM; Subkey: "Software\Classes\DoenaSoft.DVDProfiler.EnhancedFeatures.Plugin"; Flags: dontcreatekey deletekey
; Unregister
Root: HKCR; Subkey: "CLSID\{{58E2706B-250B-4ED4-BD69-82DEA5A1BBD6}"; Flags: dontcreatekey uninsdeletekey
Root: HKCR; Subkey: "DoenaSoft.DVDProfiler.EnhancedFeatures.Plugin"; Flags: dontcreatekey uninsdeletekey
Root: HKCU; Subkey: "Software\Invelos Software\DVD Profiler\Plugins\Identified"; ValueType: none; ValueName: "{{58E2706B-250B-4ED4-BD69-82DEA5A1BBD6}"; ValueData: "0"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Invelos Software\DVD Profiler_beta\Plugins\Identified"; ValueType: none; ValueName: "{{58E2706B-250B-4ED4-BD69-82DEA5A1BBD6}"; ValueData: "0"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Classes\CLSID\{{58E2706B-250B-4ED4-BD69-82DEA5A1BBD6}"; Flags: dontcreatekey uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\DoenaSoft.DVDProfiler.EnhancedFeatures.Plugin"; Flags: dontcreatekey uninsdeletekey

[Code]
function IsDotNET35Detected(): boolean;
// Function to detect dotNet framework version 2.0
// Returns true if it is available, false it's not.
var
dotNetStatus: boolean;
begin
dotNetStatus := RegKeyExists(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5');
Result := dotNetStatus;
end;

function InitializeSetup(): Boolean;
// Called at the beginning of the setup package.
begin

if not IsDotNET35Detected then
begin
MsgBox( 'The Microsoft .NET Framework version 3.5 is not installed. Please install it and try again.', mbInformation, MB_OK );
Result := false;
end
else
Result := true;
end;

