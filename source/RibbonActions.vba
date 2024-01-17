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

Private Sub cmdCasinoSettings_onAction(ByVal Control As IRibbonControl)
  ' Call CasinoActions.ExportSettings
End Sub

