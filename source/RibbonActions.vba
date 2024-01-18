Option Explicit
Option Private Module

Private Sub cmdInitialize_onAction()
    Call LotteryActions.Initialize
End Sub

Private Sub cmdGenCampConfig_onAction()
    Call LotteryActions.GenConfig
End Sub

Private Sub cmdRollDice_onAction()
    Call LotteryActions.RunLottery
End Sub

Private Sub cmdTest_onAction()
  Call LotteryActions.Test
End Sub

Private Sub cmdRemoveDuplicates_onAction()
    Call LotteryActions.RemoveDuplicates
End Sub

