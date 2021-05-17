# v2
A fully customizable toolkit designed to easily add constructed PowerShell fix actions for tier-I/ter-II execution. 

For the end user (a tier-I, tier-II, or other support technician), the tool provides a quick interface to easily search a user or computer from within ADDS, pull reverent information, and execute various actions. All of this - the shown information and its presentation, the executable actions, and even the search term themselves are all definable from within the tool itself using PowerShell script blocks.

<b>Technician view:</b>
![Alt text](web/00.png "Overview")

All settings are internally configurable. These include ADDS computer and user properties, login log definitions, standalone PowerShell tools, and more.

<b>Configuration view:</b>
![Alt text](web/01.png "Overview")

<b>Configuration view (general settings tab):</b>
![Alt text](web/02.png "Overview")

<b>Configuration view (user AD properties):</b>
![Alt text](web/03.png "Overview")

<b>Configuration view (user AD property attached script block):</b>
![Alt text](web/04.png "Overview")

All settings can be centrally published an configuration file to share internally. Clients pulling from this config will then monitor it for later updates. This can be easily accomplished be 'saving' the configuration from the configuration view - other clients that import from this will continue to monitor the published location for updates.

<b>Configuration view (exporting a configuration file):</b>
![Alt text](web/05.png "Overview")

All actions searches and script executions initiated from within the tool are logged. From within the tools view, default reports easily allow tracking of these for multiple technicians.

<b>Tool view (exporting a report):</b>
![Alt text](web/06.png "Overview")

These reports, using PSWriteHTML, will display events according to the selected timeframes.

<b>Report view:</b>
![Alt text](web/07.png "Overview")
