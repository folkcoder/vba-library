VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NetworkAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Description("Provides methods to interface with network adapters.")
'@Folder("VBALibrary.Network")
'@PredeclaredId(True)
'@Remarks("This class has its hidden VB_PredeclaredId attribute set to true in order to create a static instance of the class without needing to be instantiated.")

Option Explicit

' =============================================================================

'@Description("Determines whether the calling machine is connection to a network driver with the provided name.")
'@Param("strNetworkDriverName: The name of the network driver ")
'@Returns(""A value indicating whether the current machine is connected a network adapter matching the provided network device name.")
Public Function IsConnectedToNamedNetwork(ByVal strNetworkDriverName As String) As Boolean

Const INTERFACE_OPERATIONAL_STATUS_UP = 1

Dim objWmiService As Object
Dim vntNetworkAdapters As Variant
Dim vntNetworkAdapter As Variant

     Set objWmiService = GetObject("winmgmts:\\.\root\standardcimv2")
     Set vntNetworkAdapters = objWmiService.ExecQuery("SELECT * FROM MSFT_NetAdapter")

     For Each vntNetworkAdapter In vntNetworkAdapters
         If LCase$(vntNetworkAdapter.DriverDescription) = LCase$(strNetworkDriverName) Then
             If vntNetworkAdapter.InterfaceOperationalStatus = INTERFACE_OPERATIONAL_STATUS_UP Then
                 IsConnectedToNamedNetwork = True
             End If
         End If
     Next vntNetworkAdapter

End Function
