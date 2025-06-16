
[Setup]
AppName=Shuffle XLSX Tool
AppVersion=1.0
DefaultDirName={pf}\ShuffleXLSXTool
DefaultGroupName=Shuffle XLSX Tool
OutputBaseFilename=Setup_ShuffleRowsXlsx
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\shuffle-rows-xslx.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{desktop}\Shuffle XLSX Tool"; Filename: "{app}\shuffle-rows-xslx.exe"
