Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Interop

Class MainWindow

    Private Declare Function GetCursorPos Lib "user32" Alias "GetCursorPos" (<Out> ByRef lpPoint As POINT) As Boolean
    Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Declare Function DeleteObject Lib "gdi32" (hObject As IntPtr) As Boolean





    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    <DllImport("user32.dll")>
    Private Shared Function GetWindowLong(ByVal hwnd As IntPtr, ByVal index As Integer) As Integer
    End Function
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    <DllImport("user32.dll")>
    Private Shared Function SetWindowLong(ByVal hwnd As IntPtr, ByVal index As Integer, ByVal newStyle As Integer) As Integer
    End Function
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Structure POINT
        Public X As Integer
        Public Y As Integer
    End Structure
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Shared _x, _y As Integer
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Shared Sub SetWindowExTransparent(ByVal hwnd As IntPtr)
        Dim extendedStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
        SetWindowLong(hwnd, GWL_EXSTYLE, extendedStyle Or WS_EX_TRANSPARENT)
    End Sub
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Const WS_EX_TRANSPARENT As Integer = &H20
    Const GWL_EXSTYLE As Integer = -20
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Async Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        MyBase.OnSourceInitialized(e)
        Dim hwnd = New WindowInteropHelper(Me).Handle
        SetWindowExTransparent(hwnd)

        Await SET_MOUSE_POS()

    End Sub
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Shared Function ConvertPixelsToDIPixels(ByVal pixels As Integer, ByVal DPI As Integer) As Double

        Return CDbl(pixels) * 96 / DPI

    End Function
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Async Function SET_MOUSE_POS() As Task

        Try

            Dim point As POINT
            Me.Left = 0
            Me.Top = 0

            Dim screen As IntPtr = GetDC(0)
            Dim dpiX = GetDeviceCaps(screen, 88)
            Dim dpiY = GetDeviceCaps(screen, 88)
            Dim physicalUnitSize = 1.0R / 96.0R * dpiX
            ReleaseDC(0, screen)

            Do

                GetCursorPos(point)

                _x = ConvertPixelsToDIPixels(point.X, dpiX)
                _y = ConvertPixelsToDIPixels(point.Y, dpiY)

                Me.Left = _x - 150
                Me.Top = _y - 150
                Me.Topmost = True

                Call Magnifier_Engine(point)

                Await Task.Delay(1)

            Loop

        Catch ex As Exception


        End Try

    End Function
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Sub Magnifier_Engine(ByVal e As POINT)

        Dim ZOOM_FACTOR = 2.0

        'Dim screenLeft As Double = SystemParameters.VirtualScreenLeft
        'Dim screenTop As Double = SystemParameters.VirtualScreenTop
        'Dim screenWidth As Double = SystemParameters.VirtualScreenWidth
        'Dim screenHeight As Double = SystemParameters.VirtualScreenHeight

        'THIS IS YOUR TASK TO IMPLEMENT THE MAGNIFIER CODE FOR THE 150PX LENS
        Dim bitmap As Bitmap = New Bitmap(Convert.ToInt32(150 / ZOOM_FACTOR), Convert.ToInt32(150 / ZOOM_FACTOR))
        Dim g As Graphics = Graphics.FromImage(bitmap)
        g.CopyFromScreen(e.X, e.Y, 0, 0, bitmap.Size)

        Dispatcher.Invoke(Sub()
                              RenderImage.Source = ImageSourceFromBitmap(bitmap)
                        End Sub)



    End Sub

    Private Function ImageSourceFromBitmap(ByVal bmp As Bitmap) As ImageSource

        Dim handle = bmp.GetHbitmap()

        Try
            Return Imaging.CreateBitmapSourceFromHBitmap(handle, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
        Finally
            DeleteObject(handle)
        End Try

    End Function
End Class
