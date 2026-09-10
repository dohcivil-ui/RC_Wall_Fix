Attribute VB_Name = "modFeasibilityRecovery"
Option Explicit
' Invalid search states can move; only CheckProjectDesign can certify a solution.
Public RecoveryActive As Boolean
Public RecoveryMoves As Long
Public RecoveryProposals As Long
Public RecoveryScore As Double
Public Const RECOVERY_MODE As Integer = 2

Public Sub InitializeRecovery(d As Design, ByVal valid As Boolean)
    RecoveryActive = Not valid
    RecoveryMoves = 0: RecoveryProposals = 0: RecoveryScore = 0
    If RecoveryActive And RECOVERY_MODE > 1 Then RecoveryScore = FeasibilityMerit(d)
End Sub

Public Function AcceptRecovery(d As Design) As Boolean
    Dim proposedScore As Double
    If Not RecoveryActive Then Exit Function
    RecoveryProposals = RecoveryProposals + 1
    If RECOVERY_MODE = 1 Then
        AcceptRecovery = True
    Else
        proposedScore = FeasibilityMerit(d)
        AcceptRecovery = proposedScore <= RecoveryScore + 0.000000000001
        If RECOVERY_MODE = 3 And Not AcceptRecovery Then
            ' Declared two-percent exploration before feasibility, equal for both methods.
            AcceptRecovery = Rand(1, 100) <= 2
        End If
        If AcceptRecovery Then RecoveryScore = proposedScore
    End If
    If AcceptRecovery Then RecoveryMoves = RecoveryMoves + 1
End Function
