# SimFin Excel Plugin

Note: we are working on making the installation more user-friendly. In the meantime, follow the instructions below to install the Excel Plugin. All the Macros that will be installed can be viewed by inspecting the .bas files here on Github or once downloaded, in case you should have security concerns.

## Installation for Excel under Windows

* download all the files in this repository by clicking on the green "Code" button and then on "Download ZIP". Unzip the files once downloaded.

* Open Developer Tab in Excel (to show the developer tab, [click here for instructions](https://support.microsoft.com/en-us/office/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45))

* Click on "Visual Basic" then on "file" -> "Import File" -> Select both SimFinAPI.bas AND JsonConverter.bas one after the other.

* Click on "Tools" -> "References" -> mark "Microsoft scripting runtime" -> ok.

## Getting data from the plugin once installed

The formula is:

=SimFin(Company Ticker As String, Year As Integer, Period As String, Columname As String, API-key As String)

So for example:
=SimFin("AAPL", 2020, "FY", "Revenues", "YOUR-API-KEY")

Your API-key can be found at https://app.simfin.com/developers
