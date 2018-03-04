# WIAScanner
Scan from WIA or TWAIN compatible scanner from .NET code (C#, VB.NET, ...) with few lines of code. 

		Dim scanner As New WIAScanner()
		Dim device = scanner.SelectDevice()
		scanner.SelectDevice()
		'MsgBox(scanner.ScannerName)

		scanner.ColorMode = 1
		scanner.ResolutionDPI = 300
		scanner.ScanHeight = 4
		scanner.ScanWidth = 2.2

		Dim result As Image = scanner.Scan()
		If FacesView.Image IsNot Nothing Then FacesView.Image.Dispose()
		FacesView.Image = result

