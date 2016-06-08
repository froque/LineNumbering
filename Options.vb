Option Strict On
Option Explicit On

Imports CommandLine
Imports CommandLine.Text

Public Class Options

    <OptionAttribute("P"c, "project", Required:=True, HelpText:="Project to generate line numbers")>
    Public Property Project As String

    <OptionAttribute("O"c, "output", Required:=False, DefaultValue:="LN", HelpText:="Output directory for new source code")>
    Public Property Output As String

    <OptionAttribute("I"c, "increment", Required:=False, DefaultValue:=1, HelpText:="Line increment to use")>
    Public Property Increment As Integer

    <HelpOption>
    Public Function GetUsage() As String
        Dim h As HelpText = New HelpText()
        h.Heading = New HeadingInfo(GetType(Options).Assembly.GetName().Name, GetType(Options).Assembly.GetName().Version.ToString())
        h.AddOptions(Me)
        Return h

    End Function

End Class
