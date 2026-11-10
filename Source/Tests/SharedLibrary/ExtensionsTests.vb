#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System.Collections.Generic
Imports System.Windows.Forms

Imports DevCase.Extensions.ControlExtensions
Imports DevCase.Extensions.MenuStripExtensions
Imports DevCase.Extensions.ToolStripExtensions

Imports Microsoft.VisualStudio.TestTools.UnitTesting

#End Region

#Region " ExtensionsTests "

Namespace EasyLink.Tests.SharedLibrary

    ''' <summary>
    ''' Unit tests for the iterative (non-recursive) control and tool strip item enumeration extensions,
    ''' used to apply the visual theme to the whole user interface.
    ''' </summary>
    <STATestClass>
    Public NotInheritable Class ExtensionsTests

#Region " Helper Methods "

        ''' <summary>
        ''' Builds the following control tree:
        ''' <code>
        ''' root (Panel)
        '''  ├─ button1 (Button)
        '''  └─ inner (Panel)
        '''      ├─ button2 (Button)
        '''      └─ label (Label)
        ''' </code>
        ''' </summary>
        '''
        ''' <returns>
        ''' The root panel.
        ''' </returns>
        Private Shared Function BuildControlTree() As Panel

            Dim root As New Panel() With {.Name = "root"}
            Dim inner As New Panel() With {.Name = "inner"}

            inner.Controls.Add(New Button() With {.Name = "button2"})
            inner.Controls.Add(New Label() With {.Name = "label"})

            root.Controls.Add(New Button() With {.Name = "button1"})
            root.Controls.Add(inner)

            Return root

        End Function

        ''' <summary>
        ''' Fills the specified tool strip with the following item tree:
        ''' <code>
        ''' button (ToolStripButton)
        ''' dropDown (ToolStripDropDownButton)
        '''  └─ menu1 (ToolStripMenuItem)
        '''      └─ menu2 (ToolStripMenuItem)
        ''' </code>
        ''' </summary>
        '''
        ''' <param name="strip">
        ''' The tool strip to fill.
        ''' </param>
        Private Shared Sub FillItemTree(strip As ToolStrip)

            Dim menu1 As New ToolStripMenuItem("menu1")
            menu1.DropDownItems.Add(New ToolStripMenuItem("menu2"))

            Dim dropDown As New ToolStripDropDownButton("dropDown")
            dropDown.DropDownItems.Add(menu1)

            strip.Items.Add(New ToolStripButton("button"))
            strip.Items.Add(dropDown)

        End Sub

#End Region

#Region " ControlExtensions.ForEachControl "

        ''' <summary>
        ''' Verifies that the non-recursive iteration only visits the direct children.
        ''' </summary>
        <TestMethod>
        Public Sub ForEachControl_NonRecursive_VisitsDirectChildrenOnly()

            Using root As Panel = ExtensionsTests.BuildControlTree()
                Dim visited As New List(Of String)

                root.ForEachControl(Of Control)(False, Sub(ctrl As Control) visited.Add(ctrl.Name))

                CollectionAssert.AreEqual(New String() {"button1", "inner"}, visited)
            End Using

        End Sub

        ''' <summary>
        ''' Verifies that the recursive iteration visits every descendant, in breadth-first order.
        ''' </summary>
        <TestMethod>
        Public Sub ForEachControl_Recursive_VisitsAllDescendantsBreadthFirst()

            Using root As Panel = ExtensionsTests.BuildControlTree()
                Dim visited As New List(Of String)

                root.ForEachControl(Of Control)(True, Sub(ctrl As Control) visited.Add(ctrl.Name))

                CollectionAssert.AreEqual(New String() {"button1", "inner", "button2", "label"}, visited)
            End Using

        End Sub

        ''' <summary>
        ''' Verifies that only the controls of the specified type are passed to the action.
        ''' </summary>
        <TestMethod>
        <DataRow(False, 1)>
        <DataRow(True, 2)>
        Public Sub ForEachControl_TypeFilter_VisitsOnlyMatchingControls(recursive As Boolean, expectedCount As Integer)

            Using root As Panel = ExtensionsTests.BuildControlTree()
                Dim count As Integer = 0

                root.ForEachControl(Of Button)(recursive, Sub(btn As Button) count += 1)

                Assert.AreEqual(expectedCount, count)
            End Using

        End Sub

        ''' <summary>
        ''' Verifies that a <see langword="Nothing"/> action is rejected.
        ''' </summary>
        <TestMethod>
        Public Sub ForEachControl_NullAction_ThrowsArgumentNullException()

            Using root As Panel = ExtensionsTests.BuildControlTree()
                Assert.ThrowsExactly(Of ArgumentNullException)(Sub() root.ForEachControl(Of Control)(True, Nothing))
            End Using

        End Sub

        ''' <summary>
        ''' Verifies that tool strips are rejected, because their items must be iterated with ForEachItem instead.
        ''' </summary>
        <TestMethod>
        Public Sub ForEachControl_ToolStrip_ThrowsInvalidOperationException()

            Using strip As New ToolStrip()
                Assert.ThrowsExactly(Of InvalidOperationException)(
                    Sub() strip.ForEachControl(Of Control)(False, Sub(ctrl As Control) ctrl.Tag = Nothing))
            End Using

        End Sub

#End Region

#Region " ToolStripExtensions.ForEachItem "

        ''' <summary>
        ''' Verifies that the non-recursive iteration only visits the top-level items.
        ''' </summary>
        <TestMethod>
        Public Sub ForEachItem_NonRecursive_VisitsTopLevelItemsOnly()

            Using strip As New ToolStrip()
                ExtensionsTests.FillItemTree(strip)
                Dim visited As New List(Of String)

                strip.ForEachItem(False, Sub(item As ToolStripItem) visited.Add(item.Text))

                CollectionAssert.AreEqual(New String() {"button", "dropDown"}, visited)
            End Using

        End Sub

        ''' <summary>
        ''' Verifies that the recursive iteration visits every nested drop-down item, in breadth-first order.
        ''' </summary>
        <TestMethod>
        Public Sub ForEachItem_Recursive_VisitsAllNestedItems()

            Using strip As New ToolStrip()
                ExtensionsTests.FillItemTree(strip)
                Dim visited As New List(Of String)

                strip.ForEachItem(True, Sub(item As ToolStripItem) visited.Add(item.Text))

                CollectionAssert.AreEqual(New String() {"button", "dropDown", "menu1", "menu2"}, visited)
            End Using

        End Sub

        ''' <summary>
        ''' Verifies that only the items of the specified type are passed to the action.
        ''' </summary>
        <TestMethod>
        Public Sub ForEachItem_TypeFilter_VisitsOnlyMatchingItems()

            Using strip As New ToolStrip()
                ExtensionsTests.FillItemTree(strip)
                Dim visited As New List(Of String)

                strip.ForEachItem(Of ToolStripMenuItem)(True, Sub(item As ToolStripMenuItem) visited.Add(item.Text))

                CollectionAssert.AreEqual(New String() {"menu1", "menu2"}, visited)
            End Using

        End Sub

        ''' <summary>
        ''' Verifies that a <see langword="Nothing"/> action is rejected.
        ''' </summary>
        <TestMethod>
        Public Sub ForEachItem_NullAction_ThrowsArgumentNullException()

            Using strip As New ToolStrip()
                Dim nullAction As Action(Of ToolStripItem) = Nothing
                Assert.ThrowsExactly(Of ArgumentNullException)(Sub() strip.ForEachItem(True, nullAction))
            End Using

        End Sub

#End Region

#Region " MenuStripExtensions.ForEachItem "

        ''' <summary>
        ''' Verifies that the menu strip overload iterates the same nested items as the tool strip overload.
        ''' </summary>
        <TestMethod>
        Public Sub MenuStripForEachItem_Recursive_VisitsAllNestedItems()

            Using menu As New MenuStrip()
                Dim fileMenu As New ToolStripMenuItem("File")
                Dim recentMenu As New ToolStripMenuItem("Recent")
                recentMenu.DropDownItems.Add(New ToolStripMenuItem("Item"))
                fileMenu.DropDownItems.Add(recentMenu)
                menu.Items.Add(fileMenu)
                menu.Items.Add(New ToolStripMenuItem("Help"))

                Dim visited As New List(Of String)

                menu.ForEachItem(True, Sub(item As ToolStripItem) visited.Add(item.Text))

                CollectionAssert.AreEqual(New String() {"File", "Help", "Recent", "Item"}, visited)
            End Using

        End Sub

#End Region

    End Class

End Namespace

#End Region
