Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.VisualStudio.TestTools.UITesting

Namespace Roser.CodedUI.Common

    Class ImageLocator

        ''' <summary>
        ''' Return Point of image on screen
        ''' </summary>
        Public Shared Function OnScreen(image As Bitmap, Optional retry As Integer = 5) As Point
            Dim p As Point
            For index As Integer = 1 To retry
                p = Search(GetScreen, image)
                If p = Nothing Then
                    Threading.Thread.Sleep(1000)
                    If index = retry Then Throw New Exception("Image not found on the current screen")
                Else
                    Exit For
                End If
            Next

            ' return center of the found image
            Return New Point(p.X + CInt((image.Width / 2)), p.Y + CInt((image.Height / 2)))
        End Function


        ''' <summary>
        ''' Drag first image to the second image
        ''' </summary>
        Public Shared Sub Drag(fromImage As Bitmap, toImage As Bitmap)
            Mouse.MouseDragSpeed = 1000
            Mouse.MouseMoveSpeed = 1000
            Mouse.Move(OnScreen(fromImage))
            Mouse.StartDragging()
            Mouse.StopDragging(OnScreen(toImage))
        End Sub

        ''' <summary>
        ''' 
        ''' Finds a bitmap in a larger bitmap and returns coords
        ''' 
        ''' Original Code snippet from http://stackoverflow.com/a/12606249/1088584
        ''' By Ken Fyrstenberg https://github.com/epistemex
        ''' License CC-attribution 3.0 https://creativecommons.org/licenses/by/3.0/
        ''' </summary>
        Private Shared Function Search(src As Bitmap, ByRef bmp As Bitmap) As Point
            ' Some logic pre-checks
            If src Is Nothing OrElse bmp Is Nothing Then Return Nothing

            If src.Width = bmp.Width AndAlso src.Height = bmp.Height Then
                If src.GetPixel(0, 0) = bmp.GetPixel(0, 0) Then
                    Return New Point(0, 0)
                Else
                    Return Nothing
                End If

            ElseIf src.Width < bmp.Width OrElse src.Height < bmp.Height Then
                Return Nothing

            End If

            ' Prepare optimizations
            Dim sr As New Rectangle(0, 0, src.Width, src.Height)
            Dim br As New Rectangle(0, 0, bmp.Width, bmp.Height)

            Dim srcLock As BitmapData = src.LockBits(sr, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb)
            Dim bmpLock As BitmapData = bmp.LockBits(br, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb)

            Dim sStride As Integer = srcLock.Stride
            Dim bStride As Integer = bmpLock.Stride

            Dim srcSz As Integer = sStride * src.Height
            Dim bmpSz As Integer = bStride * bmp.Height

            Dim srcBuff(srcSz) As Byte
            Dim bmpBuff(bmpSz) As Byte

            Marshal.Copy(srcLock.Scan0, srcBuff, 0, srcSz)
            Marshal.Copy(bmpLock.Scan0, bmpBuff, 0, bmpSz)

            ' we don't need to lock the image anymore as we have a local copy
            bmp.UnlockBits(bmpLock)
            src.UnlockBits(srcLock)

            Dim x, y, x2, y2, sx, sy, bx, by, sw, sh, bw, bh As Integer
            Dim r, g, b As Byte

            Dim p As Point = Nothing

            bw = bmp.Width
            bh = bmp.Height

            sw = src.Width - bw      ' limit scan to only what we need. the extra corner
            sh = src.Height - bh     ' point we need is taken care of in the loop itself.

            ' Scan source for bitmap
            For y = 0 To sh
                sy = y * sStride
                For x = 0 To sw

                    sx = sy + x * 3
                    ' Find start point/pixel
                    r = srcBuff(sx + 2)
                    g = srcBuff(sx + 1)
                    b = srcBuff(sx)

                    If r = bmpBuff(2) AndAlso g = bmpBuff(1) AndAlso b = bmpBuff(0) Then
                        p = New Point(x, y)

                        ' We have a pixel match, check the region
                        For y2 = 0 To bh - 1
                            by = y2 * bStride
                            For x2 = 0 To bw - 1
                                bx = by + x2 * 3

                                sy = (y + y2) * sStride
                                sx = sy + (x + x2) * 3

                                r = srcBuff(sx + 2)
                                g = srcBuff(sx + 1)
                                b = srcBuff(sx)

                                If Not (r = bmpBuff(bx + 2) AndAlso
                                        g = bmpBuff(bx + 1) AndAlso
                                        b = bmpBuff(bx)) Then

                                    ' Not matching, continue checking
                                    p = Nothing
                                    sy = y * sStride
                                    Exit For
                                End If

                            Next
                            If p = Nothing Then Exit For
                        Next
                    End If 'end of region check

                    If p <> Nothing Then Exit For
                Next
                If p <> Nothing Then Exit For
            Next

            Return p

        End Function

        ''' <summary>
        ''' Returns a Bitmap of the current multihead screen desktop
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>Created because UITestControl.Desktop.CaptureImage only captures a single screen</remarks>
        Public Shared Function GetScreen() As Bitmap
            Dim image As Bitmap = New Bitmap(Screen.AllScreens.Sum(Function(s As Screen) s.Bounds.Width), Screen.AllScreens.Max(Function(s As Screen) s.Bounds.Height))
            Dim gfx As Graphics = Graphics.FromImage(image)

            ' Capture fullscreen, also multiple screens
            gfx.CopyFromScreen(SystemInformation.VirtualScreen.X, SystemInformation.VirtualScreen.Y, 0, 0, SystemInformation.VirtualScreen.Size)

            Return image
        End Function

    End Class

End Namespace
