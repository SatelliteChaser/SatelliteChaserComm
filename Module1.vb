Option Explicit On
Option Strict Off        'necessary for Paramount object

Imports ASCOM
Imports ASCOM.Astrometry
Imports ASCOM.Astrometry.AstroUtils
Imports ASCOM.DeviceInterface
Imports ASCOM.Utilities

'These are the communication functions used by Satellite Chaser, used to send commands to the mount.
'ASCOM is tested and working 
'Paramount is NOT tested and likely not working

Module Module1

    Dim mountconnected As Boolean = False 'for all connections
    Dim CommType As Integer = 0 '0 = ASCOM, 1 = Paramount

    Private objtelescope As ASCOM.DriverAccess.Telescope 'ASCOM
    Dim TelescopeName As String 'ASCOM

    Dim tsx_ts As Object 'Paramount object

    Sub Main()
        'example test for an ASCOM mount
        'CallAndConnectASCOM()
        'SlewTest()
        'Btn_Connect_Click() 'Disconnect ASCOM
    End Sub

    'ASCOM Mount Chooser Window
    Private Sub Btn_ConnectMount_Click()
        Dim obj As New ASCOM.Utilities.Chooser
        obj.DeviceType = "Telescope"
        TelescopeName = obj.Choose()
    End Sub

    'ASCOM Mount Connection
    Private Sub Btn_Connect_Click()
        If mountconnected = False Then
            CommType = 0

            objtelescope = New ASCOM.DriverAccess.Telescope(TelescopeName)
            objtelescope.Connected = True
            mountconnected = True

            If objtelescope.CanMoveAxis(TelescopeAxes.axisPrimary) = False Or objtelescope.CanMoveAxis(TelescopeAxes.axisSecondary) = False Then
                Console.WriteLine("The driver of this mount has specified that it doensn't support the MoveAxis command! Tracking might not be possible.")
            End If
        Else
            objtelescope.Connected = False
            mountconnected = False
        End If

    End Sub

    'Paramount connection
    Private Sub btn_ConnectParamount_Click()
        If mountconnected = False Then
            CommType = 1
            Try
                tsx_ts = CreateObject("TheSkyX.sky6RASCOMTele")
            Catch ex As Exception
                Console.WriteLine("Trying alternate COM object.")
                tsx_ts = CreateObject("TheSkyX.Sky6RASCOMTele") 'big S
            End Try

            'Connect to the telescope
            tsx_ts.Connect()

            'See if connection failed
            If (tsx_ts.IsConnected = 0) Then
                MsgBox("Connection failed.")
                Return
            End If

            tsx_ts.Asynchronous = 1

            mountconnected = True
        Else
            tsx_ts.Disconnect()
            mountconnected = False
        End If
    End Sub

    Private Sub SetMoveAxis(Axis As Integer, Amount As Double) 'input in deg

        Select Case CommType
            Case 0 'ASCOM
                If Axis = 0 Then
                    objtelescope.MoveAxis(TelescopeAxes.axisPrimary, Amount)
                Else
                    objtelescope.MoveAxis(TelescopeAxes.axisSecondary, Amount)
                End If

            Case 1 'PARAMOUNT
                If Axis = 0 Then
                    Dim DEC As Double = tsx_ts.dDecTrackingRate
                    tsx_ts.setTrackingRates(1, 0, Amount * 3600, DEC) 'in Arcsec per second
                Else
                    Dim RA As Double = tsx_ts.dRaTrackingRate
                    tsx_ts.setTrackingRates(1, 0, RA, Amount * 3600) 'in Arcsec per second
                End If
            Case Else
        End Select

    End Sub

    Private Function GetRightAscension() As Integer

        Select Case CommType
            Case 0 'ASCOM
                Return objtelescope.RightAscension

            Case 1 'PARAMOUNT
                tsx_ts.GetRaDec()
                Return tsx_ts.dRa / 3600 'return in deg
            Case Else
        End Select

        Return -1
    End Function

    Private Sub SetAbortSlew()

        Select Case CommType
            Case 0 'ASCOM
                objtelescope.AbortSlew()
            Case 1 'PARAMOUNT
                tsx_ts.Abort()
            Case Else
        End Select

    End Sub

    Private Function GetName() As String

        Select Case CommType
            Case 0 'ASCOM
                Return objtelescope.Name
            Case 1 'PARAMOUNT
                Return "Mount Connected via TheSkyX"
            Case Else
        End Select

        Return "Error"
    End Function

    Private Sub SetSyncToCoordinates(RA As Double, DEC As Double) 'both inputs in deg

        Select Case CommType
            Case 0 'ASCOM
                objtelescope.SyncToCoordinates(RA, DEC)
            Case 1 'PARAMOUNT
                tsx_ts.Sync(RA, DEC, "Satellite")
            Case Else
        End Select

    End Sub

    Private Sub SetTracking(DoTracking As Boolean)

        Select Case CommType
            Case 0 'ASCOM
                If DoTracking = True Then
                    objtelescope.Tracking = True
                Else
                    objtelescope.Tracking = False
                End If
            Case 1 'PARAMOUNT
                If DoTracking = True Then
                    tsx_ts.setTrackingRates(1, 1, 0, 0)
                Else
                    tsx_ts.setTrackingRates(0, 1, 0, 0)
                End If
            Case Else
        End Select

    End Sub

    Private Sub SetSlewToCoordinatesAsync(RA As Double, DEC As Double) 'both inputs in deg

        Select Case CommType
            Case 0 'ASCOM
                objtelescope.SlewToCoordinatesAsync(RA, DEC)
            Case 1 'PARAMOUNT
                tsx_ts.SlewToRaDec(RA, DEC, "Satellite")
            Case Else
        End Select

    End Sub


    'example functions for testing implementation
    'you can add new functions here for easier testing
    Sub CallAndConnectASCOM()
        Btn_ConnectMount_Click()
        Btn_Connect_Click()
    End Sub

    Sub ConnectParamount()
        btn_ConnectParamount_Click()
    End Sub

    Sub SlewTest()
        'incrementally speed up and down to 1 deg/sec
        'Satellite Chaser sends speed update 10 times per second, this test should pass for the mount
        If mountconnected Then
            For number As Double = 0 To 1 Step 0.01
                Console.WriteLine("Moving RA at " + CStr(number) + " Deg/sec")
                SetMoveAxis(0, number)
                System.Threading.Thread.Sleep(100)
            Next
            For number As Double = 1 To 0 Step -0.01
                Console.WriteLine("Moving RA at " + CStr(number) + " Deg/sec")
                SetMoveAxis(0, number)
                System.Threading.Thread.Sleep(100)
            Next
            SetAbortSlew()
        End If
    End Sub

End Module
