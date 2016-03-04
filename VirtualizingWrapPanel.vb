Imports System.Windows.Controls.Primitives
Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.Windows.Media.Animation

Public Class VirtualizingWrapPanel
    Inherits VirtualizingPanel
    Implements IScrollInfo

#Region " Private Fields"


    Private _NumberOfRows As Integer

    Private _TotalHeightOfVisibleChildren As Double

    Private _ChildSize As Size = Size.Empty
    Private _ChildPositionList As New Dictionary(Of Integer, Point)

    Private _Extent As New Size(0, 0)

    Private _Offset As New Point(0, 0)
    Private _Viewport As New Size(0, 0)

#End Region

#Region " Dependency Properties"

    Public Shared ReadOnly ItemsPerRowProperty As DependencyProperty = DependencyProperty.Register("ItemsPerRow", GetType(Integer), GetType(VirtualizingWrapPanel), New PropertyMetadata(0))

    Public Property ItemsPerRow() As Integer
        Get
            Return DirectCast(Me.GetValue(ItemsPerRowProperty), Integer)
        End Get
        Set(ByVal value As Integer)
            Me.SetValue(ItemsPerRowProperty, value)
        End Set
    End Property


#End Region

#Region " Constructor"

    Public Sub New()
        Me.RenderTransform = _trans
    End Sub

#End Region

#Region " BaseClass Overrides"

    Protected Overrides Function MeasureOverride(availableSize As Size) As Size
        If ItemsControl.GetItemsOwner(Me) Is Nothing OrElse ItemsControl.GetItemsOwner(Me).Items.Count = 0 Then
            Return availableSize
        End If

        Me._TotalHeightOfVisibleChildren = 0
        Me._NumberOfRows = 0

        Dim children As UIElementCollection = Me.InternalChildren
        Dim generator As IItemContainerGenerator = Me.ItemContainerGenerator

        Dim startPos = generator.GeneratorPositionFromIndex(Me.GetFirstVisibleIndex())
        Dim childIndex = If((startPos.Offset = 0), startPos.Index, startPos.Index + 1)

        Using generator.StartAt(startPos, GeneratorDirection.Forward, True)

            Dim currentChildInRow = 0
            Dim firstVisibleRow = Me.GetFirstVisibleRow()
            Dim currentRow = firstVisibleRow

            Do

                Dim newlyRealized As Boolean

                Dim child As UIElement = DirectCast(generator.GenerateNext(newlyRealized), UIElement)

                If child IsNot Nothing Then

                    If newlyRealized Then
                        ' Figure out if we need to insert the child at the end or somewhere in the middle
                        If childIndex >= children.Count Then
                            MyBase.AddInternalChild(child)
                        Else
                            MyBase.InsertInternalChild(childIndex, child)
                        End If
                        generator.PrepareItemContainer(child)
                    Else
                        Debug.Assert(child.Equals(Me.Children(childIndex)), "Wrong child was generated")
                    End If


                    If Me._ChildSize = Size.Empty Then
                        child.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))

                        ' All Children Will Have The Same Size (Only Measure The First Child)
                        Me._ChildSize = child.DesiredSize

                        ' Calculate The Number Of Children Per Row
                        If Me.ItemsPerRow = 0 Then
                            Me.ItemsPerRow = CInt(Math.Floor(availableSize.Width / Me._ChildSize.Width))
                        Else

                            Me._ChildSize.Width = Math.Floor(availableSize.Width / Me.ItemsPerRow)

                        End If
                    Else
                        child.Measure(_ChildSize)
                    End If


                    If currentChildInRow >= Me.ItemsPerRow Then
                        currentRow += 1
                        currentChildInRow = 0

                        Me._TotalHeightOfVisibleChildren = (currentRow - firstVisibleRow) * _ChildSize.Height

                    End If

                    If Me._TotalHeightOfVisibleChildren >= (_Viewport.Height * 2) Then
                        Exit Do
                    End If

                    If childIndex >= Me.GetChildCount - 1 Then
                        Exit Do
                    End If

                    childIndex += 1
                    currentChildInRow += 1

                Else

                    Exit Do

                End If

            Loop

            Me._NumberOfRows = CInt(Math.Ceiling(Me.GetChildCount / Me.ItemsPerRow))

        End Using

        UpdateViewPortAndExtent(availableSize)

        Return availableSize
    End Function

    Protected Overrides Function ArrangeOverride(finalSize As Size) As Size
        Dim generator As IItemContainerGenerator = Me.ItemContainerGenerator

        SyncLock Me.Children.SyncRoot

            For i As Integer = 0 To Me.Children.Count - 1

                Dim child As UIElement = Me.Children(i)
                Dim itemIndex As Integer = generator.IndexFromGeneratorPosition(New GeneratorPosition(i, 0))

                Dim row = Math.Floor(itemIndex / Me.ItemsPerRow)
                Dim column = itemIndex Mod Me.ItemsPerRow

                child.Arrange(New Rect(column * Me._ChildSize.Width, row * Me._ChildSize.Height, Me._ChildSize.Width, Me._ChildSize.Height))

            Next

        End SyncLock

        Return finalSize
    End Function

    Protected Overrides Sub BringIndexIntoView(index As Integer)
        SetFirstRowViewItemIndex(index)
    End Sub

    Protected Overrides Sub OnItemsChanged(sender As Object, args As ItemsChangedEventArgs)
        Select Case args.Action
            Case NotifyCollectionChangedAction.Remove, NotifyCollectionChangedAction.Replace, NotifyCollectionChangedAction.Move
                RemoveInternalChildRange(args.Position.Index, args.ItemUICount)
            Case NotifyCollectionChangedAction.Reset
                Me.SetVerticalOffset(0)
        End Select
    End Sub

    Protected Overrides Sub OnInitialized(e As EventArgs)
        AddHandler SizeChanged, AddressOf OnSizeChanged

        MyBase.OnInitialized(e)
    End Sub

#End Region

#Region " Private Methods"

    Private Sub UpdateViewPortAndExtent(availableSize As Size)
        If _Viewport <> availableSize Then
            _Viewport = availableSize

            If Me.ScrollOwner IsNot Nothing Then
                Me.ScrollOwner.InvalidateScrollInfo()
            End If
        End If

        Dim extentHeight = (Me._NumberOfRows * Me._ChildSize.Height)

        If extentHeight <> _Extent.Height Then
            _Extent.Height = extentHeight

            If Me.ScrollOwner IsNot Nothing Then
                Me.ScrollOwner.InvalidateScrollInfo()
            End If
        End If

    End Sub

    Private Sub OnSizeChanged(sender As Object, e As EventArgs)
        If _Viewport.Width <> 0 Then
            MeasureOverride(_Viewport)
        End If
    End Sub

    Private Sub SetFirstRowViewItemIndex(index As Integer)
        SetVerticalOffset((index) / Math.Floor((_Viewport.Width) / _ChildSize.Width))
        SetHorizontalOffset((index) / Math.Floor((_Viewport.Height) / _ChildSize.Height))
    End Sub

    Private Function GetFirstVisibleIndex() As Integer
        Return GetFirstVisibleRow() * Me.ItemsPerRow
    End Function

    Private Function GetFirstVisibleRow() As Integer
        If _ChildSize = Size.Empty Then

            Return 0

        Else

            Return CInt(Math.Ceiling(_Offset.Y / _ChildSize.Height))

        End If
    End Function

    Private Function GetChildCount() As Integer
        Dim objParentItemsControl As ItemsControl = ItemsControl.GetItemsOwner(Me)

        If objParentItemsControl.HasItems = True Then
            Return objParentItemsControl.Items.Count
        Else
            Return 0
        End If
    End Function

#End Region

#Region " IScrollInfo Implementation"

    Private m_CanVerticallyScroll As Boolean
    Private m_CanHorizontallyScroll As Boolean
    Private m_ScrollOwner As ScrollViewer

    Public Property CanHorizontallyScroll() As Boolean Implements IScrollInfo.CanHorizontallyScroll
        Get
            Return m_CanHorizontallyScroll
        End Get
        Set(value As Boolean)
            m_CanHorizontallyScroll = value
        End Set
    End Property

    Public Property CanVerticallyScroll() As Boolean Implements IScrollInfo.CanVerticallyScroll
        Get
            Return m_CanVerticallyScroll
        End Get
        Set(value As Boolean)
            m_CanVerticallyScroll = value
        End Set
    End Property

    Public ReadOnly Property ExtentHeight() As Double Implements IScrollInfo.ExtentHeight
        Get
            Return _Extent.Height
        End Get
    End Property

    Public ReadOnly Property ExtentWidth() As Double Implements IScrollInfo.ExtentWidth
        Get
            Return _Extent.Width
        End Get
    End Property

    Public ReadOnly Property HorizontalOffset() As Double Implements IScrollInfo.HorizontalOffset
        Get
            Return _Offset.X
        End Get
    End Property

    Public Property ScrollOwner() As ScrollViewer Implements IScrollInfo.ScrollOwner
        Get
            Return m_ScrollOwner
        End Get
        Set(value As ScrollViewer)
            m_ScrollOwner = value
        End Set
    End Property

    Public ReadOnly Property VerticalOffset() As Double Implements IScrollInfo.VerticalOffset
        Get
            Return _Offset.Y
        End Get
    End Property

    Public ReadOnly Property ViewportHeight() As Double Implements IScrollInfo.ViewportHeight
        Get
            Return _Viewport.Height
        End Get
    End Property

    Public ReadOnly Property ViewportWidth() As Double Implements IScrollInfo.ViewportWidth
        Get
            Return _Viewport.Width
        End Get
    End Property

    Public Sub LineUp() Implements IScrollInfo.LineUp
        Throw New InvalidOperationException()
    End Sub

    Public Sub LineDown() Implements IScrollInfo.LineDown
        Throw New InvalidOperationException()
    End Sub

    Public Sub LineLeft() Implements IScrollInfo.LineLeft
        Throw New InvalidOperationException()
    End Sub

    Public Sub LineRight() Implements IScrollInfo.LineRight
        Throw New InvalidOperationException()
    End Sub

    Public Sub PageUp() Implements IScrollInfo.PageUp
        Throw New InvalidOperationException()
    End Sub

    Public Sub PageDown() Implements IScrollInfo.PageDown
        Throw New InvalidOperationException()
    End Sub

    Public Sub PageLeft() Implements IScrollInfo.PageLeft
        Throw New InvalidOperationException()
    End Sub

    Public Sub PageRight() Implements IScrollInfo.PageRight
        Throw New InvalidOperationException()
    End Sub

    Public Sub MouseWheelUp() Implements IScrollInfo.MouseWheelUp
        SetVerticalOffset(Me.VerticalOffset - 10)
    End Sub

    Public Sub MouseWheelDown() Implements IScrollInfo.MouseWheelDown
        SetVerticalOffset(Me.VerticalOffset + 10)
    End Sub

    Public Sub MouseWheelLeft() Implements IScrollInfo.MouseWheelLeft
        Throw New InvalidOperationException()
    End Sub

    Public Sub MouseWheelRight() Implements IScrollInfo.MouseWheelRight
        Throw New InvalidOperationException()
    End Sub

    Public Function MakeVisible(visual As Visual, rectangle As Rect) As Rect Implements IScrollInfo.MakeVisible
        Return New Rect
    End Function

    Public Sub SetHorizontalOffset(offset As Double) Implements IScrollInfo.SetHorizontalOffset
        Throw New InvalidOperationException()
    End Sub

    Public Sub SetVerticalOffset(offset As Double) Implements IScrollInfo.SetVerticalOffset
        If offset < 0 OrElse _Viewport.Height >= _Extent.Height Then
            offset = 0
        Else
            If offset + _Viewport.Height >= _Extent.Height Then
                offset = _Extent.Height - _Viewport.Height
            End If
        End If

        _Offset.Y = offset

        If Me.ScrollOwner IsNot Nothing Then
            Me.ScrollOwner.InvalidateScrollInfo()
        End If

        _trans.Y = -offset

        Me.InvalidateMeasure()
    End Sub

    Private _trans As New TranslateTransform()

#End Region

End Class
