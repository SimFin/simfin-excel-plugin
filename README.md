# SimFin Excel Plugin

Follow the instructions below to install the Excel Plugin. Check out the Example.xlsx file for a detailed usage example.

## Installation for Excel under Windows

* Download all the files in this repository by clicking on the green "Code" button and then on "Download ZIP". Unzip the files once downloaded. Alternatively you can also just download the SimFinApi.bas file.

* Open Developer Tab in Excel (to show the developer tab, [click here for instructions](https://support.microsoft.com/en-us/office/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45))

* Click on "Visual Basic" then on "file" -> "Import File" -> Select SimFinAPI.bas that you downloaded from this repository.

## Getting data from the plugin once installed

The formula is:

```
=SimFin(Company Ticker As String; Year As Integer; Period As String; Columname As String; API-key As String; [optional] ttm as String = "false"; [optional] asReported as String = "false")
```

To see the inputs in the formula, type "=simfin(" and then press **Ctrl + Shift + a**

So for example:
```=SimFin("AAPL"; 2020; "FY"; "Revenue"; "YOUR-API-KEY")```

or if you want to retrieve TTM values for Q1 2020:
```=SimFin("AAPL"; 2020; "Q1"; "Revenue"; "YOUR-API-KEY"; "true")```

Your API-key can be found at https://app.simfin.com/developers

To see all the available line item names, visit https://app.simfin.com/developers/columns
