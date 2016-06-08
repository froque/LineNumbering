# LineNumbering
VB6 Line Numbering Build Tool

# Usage
```
$ bin\LineNumbering.exe
LineNumbering 1.0.0.0

  P, project      Required. Project to generate line numbers
  O, output       (Default: LN) Output directory for new source code
  I, increment    (Default: 1) Line increment to use
  help            Display this help screen.
```
# Build 
```
set PATH=%PATH%;"C:\Program Files (x86)\MSBuild\14.0\Bin\"
set PATH=%PATH%;"C:\nuget\"
nuget.exe restore
MSBuild.exe /verbosity:minimal /p:Configuration=Release LineNumbering.sln
```
# Credit
Adapted from http://www.contactandcoil.com/software/vb6-line-numbering-build-tool/
