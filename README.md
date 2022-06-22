# v2 Configurable Support Toolkit
A fully customizable, multithreaded support toolkit designed to easily add constructed PowerShell fix actions for tier-I/tier-II execution. Written in WPF/PowerShell.

Requires:<br>
[PoshRSJob](https://www.powershellgallery.com/packages/PoshRSJob/1.7.4.4)<br>
[PSWriteHTML](https://www.powershellgallery.com/packages/PSWriteHTML/0.0.148)<br>
[RSAT](https://www.microsoft.com/en-us/download/details.aspx?id=45520)<br>

For the end user (a tier-I, tier-II, or other support technician), the tool provides a quick interface to easily search a user or computer from within ADDS, pull relevant information, and execute various actions. All of this - the shown information and its presentation, the executable actions, and even the search term themselves are all definable from within the tool itself using PowerShell script blocks.

<b>Technician view:</b>
![Alt text](web/overView.gif "Overview")

All settings are internally configurable. These include ADDS computer and user properties, login log definitions, standalone PowerShell scripts, remote access clients, and more.

<b>Configuration view:</b>
<img align="center" src="web/configView.gif">


<b>Configuration view (general settings):</b>
<br>
<img align="center" src="web/02.png">

<b>Configuration view (user AD properties):</b>
<br>
<img align="center" src="web/03.png">

<b>Configuration view (user AD property attached script block):</b>
<br>
<img align="center" src="web/04.png">

All settings can be centrally published to a configuration file to share internally. Clients pulling from this config will then monitor it for later updates. This can be easily accomplished be 'saving' the configuration from the configuration view - other clients that import from this will continue to monitor the published location for updates.

<b>Configuration view (exporting a configuration file):</b>
<br>
<img align="center" src="web/05.png">

All actions, queries, and script executions initiated from within the tool are logged. From within the tools view, default reports easily allow metric tracking for all technicians using the tool.

<b>Tool view (exporting a report):</b>
<img align="center" src="web/06.png">

These reports, using PSWriteHTML, will display events according to the selected timeframes.

<b>Report view:</b>
<img align="center" src="web/07.png">
