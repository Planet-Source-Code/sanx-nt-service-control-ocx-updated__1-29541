VERSION 5.00
Begin VB.UserControl ctlService 
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   525
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   450
   ScaleWidth      =   525
   ToolboxBitmap   =   "ctlService.ctx":0000
   Windowless      =   -1  'True
End
Attribute VB_Name = "ctlService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim svcName As String

Enum ServiceStartType
    svcBootStart = &H0
    svcSystemStart = &H1
    svcAutoStart = &H2
    svcDemandStart = &H3
    svcDisabled = &H4
End Enum

Public Function StartService() As Boolean

StartService = StartSvc(svcName)

End Function

Public Function StopService() As Boolean

StopService = StopSvc(svcName)

End Function

Public Function PauseService() As Boolean

PauseService = PauseSvc(svcName)

End Function

Public Function ResumeService() As Boolean

ResumeService = ResumeSvc(svcName)

End Function

Public Function QueryService() As String

QueryService = EnumSvcState(QuerySvc(svcName))

End Function

Public Property Get ServiceName() As String

ServiceName = svcName

End Property

Public Property Let ServiceName(ByVal vSvcName As String)

svcName = vSvcName

End Property

Public Function SetStartType(serviceStart As ServiceStartType) As Boolean

SetStartType = SetSvcStartType(svcName, serviceStart)

End Function
