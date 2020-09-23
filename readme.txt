SERVICE MANAGER OCX CONTROL
===========================

The control has only one parameter that needs setting:
"ServiceName". This value is persistent and can be set
at both run- and design-time.

The control has six functions - the actions of which
should be self-evident:

StartService
StopService
PauseService
ResumeService
QueryService
SetStartType

The first four functions return a boolean value indicating
the success of the action: True for successful, False for
failure.

The QueryService function returns a string value indicating
the current status of the service queried.
The values returned are:
"Stopped"
"Starting"
"Stopping"
"Started"
"Continuing"
"Pausing"
"Paused"
"Query Not Successful"

The SetStartType function also returns a boolean value, but
requires a parameter passed to it indicating the Service Start
Type to be set. I have used an Enum-ed variable, but the actual
values are from 0 to 4 indicating:
Boot Start
SystemStart
AutoStart
DemandStart
Disabled

The most common values are the last three - Automatic, Manual and
Disabled respectively.

Any queries, bugs or gripes:
dan@sanx.org