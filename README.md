# simfin-excel-plugin

* download all the files in this repository by clicking on the green "Code" button and then on "Download ZIP". Unzip the files once downloaded.

## Setup for Excel under Windows
* Open Developer Tab in Excel (to show the developer tab, [click here for instructions](https://support.microsoft.com/en-us/office/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45))

* Click on "Visual Basic" then on "file" - "Import File" - Select both SimFinAPI.bas AND JsonConverter.bas one after the other.
* Click on "Tools" -> "References" -> mark "Microsoft scripting runtime" -> ok

The formula is:

=SimFin(Ticker As String, Year As String, Period As String, Columname As String, API-key As String)

see the example.xlsx workbook for details
