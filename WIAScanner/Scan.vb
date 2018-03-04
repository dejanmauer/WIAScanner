Imports WIA
Imports System.Runtime.InteropServices
Imports System.Collections
Imports System.IO
Imports System.Drawing

Public Class Scan

	Private pColorMode As Integer = 1
	Public Property ColorMode() As Integer
		Get
			Return pColorMode
		End Get
		Set(ByVal value As Integer)
			'(gray is 2, color 1, bw 4, unspecified 0
			pColorMode = value
		End Set
	End Property

	Private pScannerName As String
	Public ReadOnly Property ScannerName() As String
		Get
			Return pScannerName
		End Get
	End Property

	Private pResolutionDPI As Integer
	Public Property ResolutionDPI() As Integer
		Get
			Return pResolutionDPI
		End Get
		Set(ByVal value As Integer)
			pResolutionDPI = value
		End Set
	End Property

	' Surface in INCHES 
	Private pScanTop As Double
	Public Property ScanTop() As Double
		Get
			Return pScanTop
		End Get
		Set(ByVal value As Double)
			pScanTop = value
		End Set
	End Property


	Private pScanLeft As Double
	Public Property ScanLeft() As Double
		Get
			Return pScanLeft
		End Get
		Set(ByVal value As Double)
			pScanLeft = value
		End Set
	End Property


	Private pScanWidth As Double
	Public Property ScanWidth() As Double
		Get
			Return pScanWidth
		End Get
		Set(ByVal value As Double)
			pScanWidth = value
		End Set
	End Property


	Private pScanHeight As Double
	Public Property ScanHeight() As Double
		Get
			Return pScanHeight
		End Get
		Set(ByVal value As Double)
			pScanHeight = value
		End Set
	End Property

	Private SelectedDevice As WIA.Device = Nothing

	Public Function SelectDevice(ByVal DeviceName As String) As WIA.Device
		Dim result As WIA.Device = Nothing

		Dim deviceManager As New WIA.DeviceManager
		Dim deviceInfos As DeviceInfos = deviceManager.DeviceInfos
		Dim deviceInfo As DeviceInfo

		For Each deviceInfo In deviceInfos
			Dim deviceProperties As WIA.Properties = deviceInfo.Properties
			Dim deviceProperty As WIA.Property

			For Each deviceProperty In deviceProperties
				Select Case deviceProperty.Name
					Case "Name"
						If deviceProperty.Value.ToString.ToLower = DeviceName.ToLower Then
							result = deviceInfo.Connect
							pScannerName = deviceProperty.Value.ToString
						End If
				End Select
			Next
		Next

		SelectedDevice = result
		Return result
	End Function

	Public Function SelectDevice() As WIA.Device
		Dim result As WIA.Device = Nothing

		Dim MyDevice As WIA.Device
		Dim MyDialog As New WIA.CommonDialogClass
		Try
			'shows selectdevice dialog, if only one device, It automatically selects the device
			'(not tested with two or more devices)
			'**Note - Device Type checks for VideoDeviceType, for webcams, in this sample
			MyDevice = MyDialog.ShowSelectDevice(WIA.WiaDeviceType.UnspecifiedDeviceType, False, True)

			If Not MyDevice Is Nothing Then
				'loops through device properties, only gets the ones we want to display
				For Each prop As WIA.Property In MyDevice.Properties
					Select Case prop.Name
						Case "Name"
							pScannerName = prop.Value.ToString
					End Select
				Next
				'sets MyDevice form level selected device
				result = MyDevice
			Else
				result = Nothing
			End If

		Catch ex As System.Exception
			MsgBox("No device available.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "No device!")
		End Try

		SelectedDevice = result
		Return result
	End Function

	Public Function GetDeviceName(ByVal device As WIA.Device) As String
		Dim result As String = ""

		Return result
	End Function

	Public Function Scan(Optional ByVal Device As WIA.Device = Nothing) As Image
		If Device IsNot Nothing Then SelectedDevice = Device

		Dim WiaObj As WIA.CommonDialog
		Dim WiaItm As WIA.Item = Nothing
		Dim WiaImg As New WIA.ImageFile
		Dim WiaDev As WIA.Device = Nothing
		Dim WiaVector As WIA.Vector = Nothing

		WiaObj = New WIA.CommonDialog

		If SelectedDevice Is Nothing Then     ' WiaDev is defined globally
			SelectedDevice = WiaObj.ShowSelectDevice(WIA.WiaDeviceType.ScannerDeviceType, False, False)
		End If

		Dim Itm As WIA.Item = Nothing

		Dim ItmProp As WIA.Property

		For Each Itm In SelectedDevice.Items
			For Each ItmProp In Itm.Properties
				Select Case ItmProp.PropertyID

					Case 6146   'Current Intent whether Color, Grayscale, Black-White
						ItmProp.Value = Me.ColorMode

					Case 6147   ' Horizontal Resolution
						ItmProp.Value = Me.ResolutionDPI

					Case 6148   ' Vertical Resolution
						ItmProp.Value = Me.ResolutionDPI

					Case 6149   ' Horizontal Starting Position (Scanning Area)
						ItmProp.Value = Me.ScanLeft * Me.ResolutionDPI

					Case 6150   ' Vertical Starting Position (Scanning Area)
						ItmProp.Value = Me.ScanTop * Me.ResolutionDPI

					Case 6151   ' Horizontal Extent (Scanning Area)
						ItmProp.Value = Me.ScanWidth * Me.ResolutionDPI

					Case 6152   ' Vertical Extent (Scanning Area)
						ItmProp.Value = Me.ScanHeight * Me.ResolutionDPI

					Case 6146   ' Current Intent
						ItmProp.Value = 4     ' Text or Line Art
				End Select
			Next
		Next

		WiaImg = WiaObj.ShowTransfer(Itm)
		WiaVector = WiaImg.FileData

		Dim b() As Byte = WiaVector.BinaryData
		Dim stream As New MemoryStream(b)
		Return Image.FromStream(stream)
	End Function


End Class
