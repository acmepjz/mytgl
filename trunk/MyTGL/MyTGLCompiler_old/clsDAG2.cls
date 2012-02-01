VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAG2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'DAG (directed acyclic graph) algorithm support class

'... TODO:
'topological sorting: easy
'strongly connected components: Tarjan's Algorithm

Private Type typeGraphNode
 tmp As Long
 'temp or flags
 '1=sorted
 '2=marked as parent
 '4=marked as children
 nDegreeIn As Long
 nDegreeOut As Long
 nIn() As Long '1-based index array
 nOut() As Long '1-based index array
End Type

Private m_tNodes() As typeGraphNode '1-based
Private m_nNodeCount As Long, m_nSortedNodeCount As Long

Private m_nSortedNode() As Long

'L <-- Empty list that will contain the sorted elements
'S <-- Set of all nodes with no incoming edges
'while S is non-empty do
'    remove a node n from S
'    insert n into L
'    for each node m with an edge e from n to m do
'        remove edge e from the graph
'        if m has no other incoming edges then
'            insert m into S
'if graph has edges then
'    output error message (graph has at least one cycle)
'else
'    output message (proposed topologically sorted order: L)
Friend Function RunTopologicalSort() As Boolean
Dim i As Long, j As Long
Dim n As Long, m As Long
Dim tmp As Long
Dim nEmptyNodes() As Long 'stack, 1-based
Dim nEmptyNodeCount As Long
Dim nSortedNodeCount As Long
'///
If m_nNodeCount <= 0 Then
 RunTopologicalSort = True
 Exit Function
End If
'///
ReDim m_nSortedNode(1 To m_nNodeCount)
ReDim nEmptyNodes(1 To m_nNodeCount)
'///
For i = 1 To m_nNodeCount
 Debug.Assert m_tNodes(i).nDegreeIn >= 0
 j = m_tNodes(i).nDegreeIn
 m_tNodes(i).tmp = j
 If j = 0 Then
  nEmptyNodeCount = nEmptyNodeCount + 1
  nEmptyNodes(nEmptyNodeCount) = i
 End If
Next i
'///
Do While nEmptyNodeCount > 0
 n = nEmptyNodes(nEmptyNodeCount)
 nEmptyNodeCount = nEmptyNodeCount - 1
 '///
 nSortedNodeCount = nSortedNodeCount + 1
 m_nSortedNode(nSortedNodeCount) = n
 '///
 For j = 1 To m_tNodes(n).nDegreeOut
  m = m_tNodes(n).nOut(j)
  tmp = m_tNodes(m).tmp - 1
  m_tNodes(m).tmp = tmp
  Debug.Assert tmp >= 0
  If tmp = 0 Then
   nEmptyNodeCount = nEmptyNodeCount + 1
   nEmptyNodes(nEmptyNodeCount) = m
  End If
 Next j
Loop
'///
For i = 1 To m_nNodeCount
 m_tNodes(i).tmp = (m_tNodes(i).tmp = 0) And 1&
Next i
'///
m_nSortedNodeCount = nSortedNodeCount
RunTopologicalSort = nSortedNodeCount = m_nNodeCount
End Function

Friend Sub Clear()
Erase m_tNodes, m_nSortedNode
m_nNodeCount = 0
m_nSortedNodeCount = 0
End Sub

Friend Sub AddEdge(ByVal objSrc As Long, ByVal objDest As Long)
On Error Resume Next
Dim s As String
Dim idxs As Long, idxe As Long
Dim tmp As Long
'///
idxs = objSrc
If idxs > m_nNodeCount Then
 m_nNodeCount = idxs
 ReDim Preserve m_tNodes(1 To m_nNodeCount)
 '///
 idxs = m_nNodeCount
End If
'///
idxe = objDest
If idxe > m_nNodeCount Then
 m_nNodeCount = idxe
 ReDim Preserve m_tNodes(1 To m_nNodeCount)
 '///
 idxe = m_nNodeCount
End If
'///
tmp = m_tNodes(idxs).nDegreeOut + 1
m_tNodes(idxs).nDegreeOut = tmp
ReDim Preserve m_tNodes(idxs).nOut(1 To tmp)
m_tNodes(idxs).nOut(tmp) = idxe
'///
tmp = m_tNodes(idxe).nDegreeIn + 1
m_tNodes(idxe).nDegreeIn = tmp
ReDim Preserve m_tNodes(idxe).nIn(1 To tmp)
m_tNodes(idxe).nIn(tmp) = idxs
End Sub

Friend Property Get NodeCount() As Long
NodeCount = m_nNodeCount
End Property

Friend Property Let NodeCount(ByVal n As Long)
If n > 0 Then
 ReDim m_tNodes(1 To n)
 ReDim m_nSortedNode(1 To n)
 m_nNodeCount = n
 m_nSortedNodeCount = 0
End If
End Property

Friend Property Get SortedNodeCount() As Long
SortedNodeCount = m_nSortedNodeCount
End Property

Friend Function SortedNode(ByVal nIndex As Long) As Long
SortedNode = m_nSortedNode(nIndex)
End Function

Friend Function IsNodeSorted(ByVal nIndex As Long) As Boolean
IsNodeSorted = (m_tNodes(nIndex).tmp And 1&) <> 0 And m_nSortedNodeCount > 0
End Function

Friend Function IsNodeMarked(ByVal nIndex As Long) As Boolean
IsNodeMarked = m_tNodes(nIndex).tmp And 6&
End Function

Friend Sub MarkNodeAndParent(ByVal nIndex As Long)
Dim nStack() As Long
Dim nStackCount As Long
Dim i As Long, j As Long
'///
If m_nNodeCount <= 0 Then Exit Sub
If nIndex <= 0 Or nIndex > m_nNodeCount Then Exit Sub
If m_tNodes(nIndex).tmp And 2& Then Exit Sub
'///
ReDim nStack(1 To m_nNodeCount)
nStackCount = 1
nStack(1) = nIndex
m_tNodes(nIndex).tmp = m_tNodes(nIndex).tmp Or 2&
'///
Do While nStackCount > 0
 nIndex = nStack(nStackCount)
 nStackCount = nStackCount - 1
 '///
 For i = 1 To m_tNodes(nIndex).nDegreeIn
  j = m_tNodes(nIndex).nIn(i)
  If m_tNodes(j).tmp And 2& Then
  Else
   nStackCount = nStackCount + 1
   nStack(nStackCount) = j
   m_tNodes(j).tmp = m_tNodes(j).tmp Or 2&
  End If
 Next i
Loop
'///
End Sub

Friend Sub MarkNodeAndChildren(ByVal nIndex As Long)
Dim nStack() As Long
Dim nStackCount As Long
Dim i As Long, j As Long
'///
If m_nNodeCount <= 0 Then Exit Sub
If nIndex <= 0 Or nIndex > m_nNodeCount Then Exit Sub
If m_tNodes(nIndex).tmp And 4& Then Exit Sub
'///
ReDim nStack(1 To m_nNodeCount)
nStackCount = 1
nStack(1) = nIndex
m_tNodes(nIndex).tmp = m_tNodes(nIndex).tmp Or 4&
'///
Do While nStackCount > 0
 nIndex = nStack(nStackCount)
 nStackCount = nStackCount - 1
 '///
 For i = 1 To m_tNodes(nIndex).nDegreeOut
  j = m_tNodes(nIndex).nOut(i)
  If m_tNodes(j).tmp And 4& Then
  Else
   nStackCount = nStackCount + 1
   nStack(nStackCount) = j
   m_tNodes(j).tmp = m_tNodes(j).tmp Or 4&
  End If
 Next i
Loop
'///
End Sub

Friend Property Get DegreeIn(ByVal nIndex As Long) As Long
DegreeIn = m_tNodes(nIndex).nDegreeIn
End Property

Friend Property Get DegreeOut(ByVal nIndex As Long) As Long
DegreeOut = m_tNodes(nIndex).nDegreeOut
End Property

Friend Property Get InputNode(ByVal nIndex As Long, ByVal nIndex2 As Long) As Long
InputNode = m_tNodes(nIndex).nIn(nIndex2)
End Property

Friend Property Get OutputNode(ByVal nIndex As Long, ByVal nIndex2 As Long) As Long
OutputNode = m_tNodes(nIndex).nOut(nIndex2)
End Property

