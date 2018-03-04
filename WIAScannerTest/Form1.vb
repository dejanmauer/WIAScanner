Public Class Form1
	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		Dim scanner As New WIAScanner.Scan()
		Dim device = scanner.SelectDevice()
		scanner.SelectDevice()
		'MsgBox(scanner.ScannerName)

		scanner.ColorMode = 1
		scanner.ResolutionDPI = 300
		scanner.ScanHeight = 3
		scanner.ScanWidth = 2.2

		Dim result As Image = scanner.Scan()
		PictureBox1.Image = result
	End Sub
End Class