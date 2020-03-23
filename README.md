# SolutionEnvironment
A Visual Studio extension that provides support for customized global environment settings on a per solution basis

## Introduction

The Solution Environment extension uses a new file called SolutionName.slnenv residing in the same directory as the solution to provide build environment variables tailored to a given solution file (.slnenv stands for "solution environment").
The Solution Environment extension executes this file at solution open time, before the start of each build, and before starting the debugger, resetting the build's environment variables accordingly.

## Usage
Inside the SolutionName.slnenv file, with one entry per line, are environmentvariablename=value entries.

```
MYPATH=c:\Src\MyDirectory\Include
EXTRA_OPTS=/D MY_DEFINE
```

The solution's path is available via a variable called `$(SolutionDir)`. The solution directory does not contain a terminating backslash.
```
MYPATH=$(SolutionDir)\Include
```

The solution's name is available through `$(SolutionName)`.
```
SOLNAME=$(SolutionName)\Include
```

Environment variables may be inserted using the `$(EnvironmentVariableName)` syntax. This has the same functionality as a batch file's `%EnvironmentVariableName%` substitution syntax.
```
PATH=$(PATH);c:\Tools;$(MYPATH)\..\Bin
```

Simple registry entries may be accessed via `%(HKLM\Path\Key)` or `%(HKCU\Path\Key)`, where `HKLM` accesses `HKEY_LOCAL_MACHINE`
and `HKCU` accesses `HKEY_CURRENT_USER`. Only string values may be retrieved.
```
MYPATH=%(HKLM\Software\MySoftware\Path)
```

An environment variable may be applied to a specific Solution Configuration. The syntax for this is `ConfigurationName:Name=Value`.
```
Debug:PATH=$(PATH);%(HKLM\Software\MySoftware\DebugPath)
Release:PATH=$(PATH);%(HKLM\Software\MySoftware\DebugPath)
```

Other .slnenv files may be included using the `include` or `forceinclude` keywords. The filename following each keyword should not contain the .slnenv extension.
```
include $(HOMEDRIVE)$(HOMEPATH)\MyPersonalDefinitions<br>
forceinclude ..\..\MandatoryDefinitions
```

Comments are specified by using -- at the beginning of the line.
```
-- This is a comment.
```

## Acknowledgments

This extension is based on Joshua Jensen's [Auto Build Environment Add-in for Visual Studio .NET](https://www.codeproject.com/Articles/3218/Auto-Build-Environment-Add-in-for-Visual-Studio-NE)