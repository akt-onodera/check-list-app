#requires -version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$script:IsBulkUpdateInProgress = $false

$defaultFormState = @{
  FirstTabDropdown1         = ""
  FirstTabDropdown2         = ""
  FirstTabDropdown3         = ""
  FirstTabDropdown4         = ""
  FirstTabTextBox1          = ""
  FirstTabDropdown5         = ""
  FirstTabTextBox2          = ""
  FirstTabTextBox3          = ""

  SecondTabDropdown1Special = ""
  SecondTabTextBoxSpecial1  = ""

  SecondTabDropdown2        = ""
  SecondTabDropdown3        = ""
  SecondTabDropdown4        = ""
  SecondTabDropdown5        = ""
  SecondTabTextBox1         = ""
  SecondTabDropdown6        = ""
  SecondTabDropdown7        = ""
  SecondTabDropdown8        = ""
  SecondTabDropdown9        = ""
  SecondTabDropdown10       = ""
}

$script:currentFormState = @{}
foreach ($stateKey in $defaultFormState.Keys) { $script:currentFormState[$stateKey] = $defaultFormState[$stateKey] }

function Get-ComboBoxSelectedText {
  param([Parameter(Mandatory)] [System.Windows.Controls.ComboBox] $comboBox)

  if ($comboBox.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) {
    return [string]$comboBox.SelectedItem.Content
  }
  return ""
}

function Set-ComboBoxSelectedByText {
  param(
    [Parameter(Mandatory)] [System.Windows.Controls.ComboBox] $comboBox,
    [AllowEmptyString()] [string] $text = ""
  )

  if ([string]::IsNullOrEmpty($text)) {
    $comboBox.SelectedIndex = -1
    $comboBox.SelectedItem = $null
    return
  }

  foreach ($item in $comboBox.Items) {
    if ($item -is [System.Windows.Controls.ComboBoxItem]) {
      if ([string]$item.Content -eq $text) {
        $comboBox.SelectedItem = $item
        return
      }
    }
  }

  $comboBox.SelectedIndex = -1
  $comboBox.SelectedItem = $null
}

function Set-ComboBoxItems {
  param(
    [Parameter(Mandatory)] [System.Windows.Controls.ComboBox] $comboBox,
    [Parameter(Mandatory)] $items,
    [AllowEmptyString()] [string] $preserveSelectedText = ""
  )

  $itemTexts = @()
  if ($items -is [string]) {
    $itemTexts = @([string]$items)
  }
  elseif ($items -is [System.Collections.IEnumerable]) {
    foreach ($value in $items) { $itemTexts += [string]$value }
  }
  else {
    $itemTexts = @([string]$items)
  }

  $comboBox.Items.Clear()
  foreach ($text in $itemTexts) {
    $comboBoxItem = New-Object System.Windows.Controls.ComboBoxItem
    $comboBoxItem.Content = $text
    $comboBox.Items.Add($comboBoxItem) | Out-Null
  }

  Set-ComboBoxSelectedByText -comboBox $comboBox -text $preserveSelectedText
}

function Focus-NextInputControl {
  param([Parameter(Mandatory)] [System.Windows.DependencyObject] $fromControl)

  $traversalRequest = New-Object System.Windows.Input.TraversalRequest([System.Windows.Input.FocusNavigationDirection]::Next)
  $null = $fromControl.MoveFocus($traversalRequest)
}

function Attach-EnterKeyToMoveNext {
  param([Parameter(Mandatory)] $control)

  if ($control -is [System.Windows.Controls.TextBox]) {
    $control.AcceptsReturn = $false
    $control.Add_PreviewKeyDown({
        if ($_.Key -eq [System.Windows.Input.Key]::Enter) {
          $_.Handled = $true
          Focus-NextInputControl -fromControl $this
        }
      }) | Out-Null
    return
  }

  if ($control -is [System.Windows.Controls.ComboBox]) {
    $control.Add_PreviewKeyDown({
        if ($_.Key -eq [System.Windows.Input.Key]::Enter -and -not $this.IsDropDownOpen) {
          $_.Handled = $true
          Focus-NextInputControl -fromControl $this
        }
      }) | Out-Null
    return
  }
}

function Set-WindowBottomLeft {
  param([Parameter(Mandatory)] [System.Windows.Window] $window)

  try {
    $workArea = [System.Windows.SystemParameters]::WorkArea
    $window.Left = $workArea.Left + 10
    $window.Top = $workArea.Bottom - $window.ActualHeight - 10
  }
  catch {}
}

function Format-ElapsedTime {
  param([Parameter(Mandatory)] [TimeSpan] $elapsedTime)
  "{0:00}:{1:00}:{2:00}" -f [int]$elapsedTime.TotalHours, $elapsedTime.Minutes, $elapsedTime.Seconds
}

$formFieldDefinitions = @(
  @{ ControlName = "FirstTabDropdown1"; StateKey = "FirstTabDropdown1"; ControlType = "ComboBox" }
  @{ ControlName = "FirstTabDropdown2"; StateKey = "FirstTabDropdown2"; ControlType = "ComboBox" }
  @{ ControlName = "FirstTabDropdown3"; StateKey = "FirstTabDropdown3"; ControlType = "ComboBox" }
  @{ ControlName = "FirstTabDropdown4"; StateKey = "FirstTabDropdown4"; ControlType = "ComboBox" }
  @{ ControlName = "FirstTabTextBox1"; StateKey = "FirstTabTextBox1"; ControlType = "TextBox" }
  @{ ControlName = "FirstTabDropdown5"; StateKey = "FirstTabDropdown5"; ControlType = "ComboBox" }
  @{ ControlName = "FirstTabTextBox2"; StateKey = "FirstTabTextBox2"; ControlType = "TextBox" }
  @{ ControlName = "FirstTabTextBox3"; StateKey = "FirstTabTextBox3"; ControlType = "TextBox" }

  @{ ControlName = "SecondTabDropdown1Special"; StateKey = "SecondTabDropdown1Special"; ControlType = "ComboBox" }
  @{ ControlName = "SecondTabTextBoxSpecial1"; StateKey = "SecondTabTextBoxSpecial1"; ControlType = "TextBox" }

  @{ ControlName = "SecondTabDropdown2"; StateKey = "SecondTabDropdown2"; ControlType = "ComboBox" }
  @{ ControlName = "SecondTabDropdown3"; StateKey = "SecondTabDropdown3"; ControlType = "ComboBox" }
  @{ ControlName = "SecondTabDropdown4"; StateKey = "SecondTabDropdown4"; ControlType = "ComboBox" }
  @{ ControlName = "SecondTabDropdown5"; StateKey = "SecondTabDropdown5"; ControlType = "ComboBox" }
  @{ ControlName = "SecondTabTextBox1"; StateKey = "SecondTabTextBox1"; ControlType = "TextBox" }
  @{ ControlName = "SecondTabDropdown6"; StateKey = "SecondTabDropdown6"; ControlType = "ComboBox" }
  @{ ControlName = "SecondTabDropdown7"; StateKey = "SecondTabDropdown7"; ControlType = "ComboBox" }

  @{ ControlName = "SecondTabDropdown8"; StateKey = "SecondTabDropdown8"; ControlType = "ComboBox" }
  @{ ControlName = "SecondTabDropdown9"; StateKey = "SecondTabDropdown9"; ControlType = "ComboBox" }
  @{ ControlName = "SecondTabDropdown10"; StateKey = "SecondTabDropdown10"; ControlType = "ComboBox" }
)

$comboBoxOptions = @{
  FirstTabDropdown1         = @("選択肢A", "選択肢B")
  FirstTabDropdown2         = @("選択肢A", "選択肢B")
  FirstTabDropdown3         = @("選択肢A", "選択肢B")
  FirstTabDropdown4         = @("選択肢A", "選択肢B")
  FirstTabDropdown5         = @("選択肢A", "選択肢B")

  SecondTabDropdown1Special = @("選択肢A", "選択肢B", "選択肢C")

  SecondTabDropdown2        = @("選択肢A", "選択肢B")
  SecondTabDropdown3        = @("選択肢A", "選択肢B")
  SecondTabDropdown4        = @("選択肢A", "選択肢B")
  SecondTabDropdown5        = @("選択肢A", "選択肢B")
  SecondTabDropdown6        = @("選択肢A", "選択肢B")
  SecondTabDropdown7        = @("選択肢A", "選択肢B")

  SecondTabDropdown8        = @("選択肢A", "選択肢B")
  SecondTabDropdown9        = @("選択肢A", "選択肢B")
  SecondTabDropdown10       = @("選択肢A", "選択肢B")
}

$windowXamlPath = Join-Path -Path $PSScriptRoot -ChildPath "MainWindow.xaml"
$windowXamlText = Get-Content -Path $windowXamlPath -Raw -Encoding UTF8
$xmlReader = New-Object System.Xml.XmlNodeReader ([xml]$windowXamlText)
$mainWindow = [Windows.Markup.XamlReader]::Load($xmlReader)

function Update-SpecialAddressTextBoxVisibility {
  param([Parameter(Mandatory)] [System.Windows.Window] $window)

  $specialDropdown = [System.Windows.Controls.ComboBox]$window.FindName("SecondTabDropdown1Special")
  $specialTextBox = [System.Windows.Controls.TextBox]$window.FindName("SecondTabTextBoxSpecial1")
  $selectedText = Get-ComboBoxSelectedText -comboBox $specialDropdown

  if ($selectedText -eq "選択肢C") {
    $specialTextBox.Visibility = "Visible"
    $window.Dispatcher.BeginInvoke([Action] { try { $specialTextBox.Focus() | Out-Null; $specialTextBox.SelectAll() } catch {} }) | Out-Null
  }
  else {
    $specialTextBox.Visibility = "Collapsed"
    if ($specialTextBox.Text -ne "") { $specialTextBox.Text = "" }
  }
}

function Apply-ComboBoxOptionsToWindow {
  param([Parameter(Mandatory)] [System.Windows.Window] $window)

  foreach ($controlName in $comboBoxOptions.Keys) {
    $comboBox = [System.Windows.Controls.ComboBox]$window.FindName($controlName)
    if ($comboBox) {
      $preservedText = Get-ComboBoxSelectedText -comboBox $comboBox
      Set-ComboBoxItems -comboBox $comboBox -items @($comboBoxOptions[$controlName]) -preserveSelectedText $preservedText
    }
  }
}

function Sync-UserInterfaceFromState {
  param([Parameter(Mandatory)] [System.Windows.Window] $window)

  $script:IsBulkUpdateInProgress = $true
  try {
    Apply-ComboBoxOptionsToWindow -window $window
    foreach ($definition in $formFieldDefinitions) {
      $control = $window.FindName($definition.ControlName)
      $value = [string]$script:currentFormState[$definition.StateKey]
      if ($definition.ControlType -eq "ComboBox") {
        Set-ComboBoxSelectedByText -comboBox $control -text $value
      }
      else {
        $control.Text = $value
      }
    }
    Update-SpecialAddressTextBoxVisibility -window $window
  }
  finally {
    $script:IsBulkUpdateInProgress = $false
  }
}

function Sync-StateFromUserInterface {
  param([Parameter(Mandatory)] [System.Windows.Window] $window)

  if ($script:IsBulkUpdateInProgress) { return }

  foreach ($definition in $formFieldDefinitions) {
    $control = $window.FindName($definition.ControlName)
    if ($definition.ControlType -eq "ComboBox") {
      $script:currentFormState[$definition.StateKey] = Get-ComboBoxSelectedText -comboBox $control
    }
    else {
      $script:currentFormState[$definition.StateKey] = $control.Text
    }
  }

  if ($script:currentFormState["SecondTabDropdown1Special"] -ne "選択肢C") {
    $script:currentFormState["SecondTabTextBoxSpecial1"] = ""
  }
}

function Clear-AllInputs {
  param([Parameter(Mandatory)] [System.Windows.Window] $window)

  $script:IsBulkUpdateInProgress = $true
  try {
    foreach ($key in $defaultFormState.Keys) { $script:currentFormState[$key] = "" }

    foreach ($definition in $formFieldDefinitions) {
      $control = $window.FindName($definition.ControlName)
      if ($definition.ControlType -eq "ComboBox") {
        $control.SelectedIndex = -1
        $control.SelectedItem = $null
      }
      else {
        $control.Text = ""
      }
    }
    Update-SpecialAddressTextBoxVisibility -window $window
  }
  finally {
    $script:IsBulkUpdateInProgress = $false
  }
}

$stopwatch = New-Object System.Diagnostics.Stopwatch
$stopwatchTimer = New-Object System.Windows.Threading.DispatcherTimer
$stopwatchTimer.Interval = [TimeSpan]::FromMilliseconds(200)

$stopwatchTimeTextBlock = [System.Windows.Controls.TextBlock]$mainWindow.FindName("StopwatchTimeTextBlock")
$startStopButton = [System.Windows.Controls.Button]$mainWindow.FindName("StartStopButton")
$resetButton = [System.Windows.Controls.Button]$mainWindow.FindName("ResetButton")
$clearButton = [System.Windows.Controls.Button]$mainWindow.FindName("ClearButton")

$statusChipContainer = [System.Windows.Controls.Border]$mainWindow.FindName("StatusChipContainer")
$statusChipTextBlock = [System.Windows.Controls.TextBlock]$mainWindow.FindName("StatusChipTextBlock")
$mainTabControl = [System.Windows.Controls.TabControl]$mainWindow.FindName("MainTabControl")

$hamburgerButton = [System.Windows.Controls.Button]$mainWindow.FindName("HamburgerButton")
$drawerOverlay = [System.Windows.Controls.Grid]$mainWindow.FindName("DrawerOverlay")
$drawerPanel = [System.Windows.Controls.Border]$mainWindow.FindName("DrawerPanel")
$drawerTranslateTransform = [System.Windows.Media.TranslateTransform]$mainWindow.FindName("DrawerTranslateTransform")

$drawerCloseButton = [System.Windows.Controls.Button]$mainWindow.FindName("DrawerCloseButton")
$drawerItem1Button = [System.Windows.Controls.Button]$mainWindow.FindName("DrawerItem1Button")
$drawerItem2Button = [System.Windows.Controls.Button]$mainWindow.FindName("DrawerItem2Button")
$drawerItem3Button = [System.Windows.Controls.Button]$mainWindow.FindName("DrawerItem3Button")

$headerTitleTextBlock = [System.Windows.Controls.TextBlock]$mainWindow.FindName("HeaderTitleTextBlock")
$item1Screen = [System.Windows.Controls.Grid]$mainWindow.FindName("Item1Screen")
$placeholderScreen = [System.Windows.Controls.Grid]$mainWindow.FindName("PlaceholderScreen")
$placeholderTextBlock = [System.Windows.Controls.TextBlock]$mainWindow.FindName("PlaceholderTextBlock")

function Set-StatusChipState {
  param(
    [Parameter(Mandatory)]
    [ValidateSet("Stopped", "Working", "Warning", "Danger")]
    [string] $state
  )

  switch ($state) {
    "Stopped" {
      $statusChipTextBlock.Text = "待機中"
      $statusChipContainer.Background = $mainWindow.FindResource("ChipStoppedBackgroundBrush")
      $statusChipTextBlock.Foreground = $mainWindow.FindResource("ChipStoppedForegroundBrush")
    }
    "Working" {
      $statusChipTextBlock.Text = "対応中"
      $statusChipContainer.Background = $mainWindow.FindResource("ChipBlueBackgroundBrush")
      $statusChipTextBlock.Foreground = $mainWindow.FindResource("ChipBlueForegroundBrush")
    }
    "Warning" {
      $statusChipTextBlock.Text = "保留延伸"
      $statusChipContainer.Background = $mainWindow.FindResource("ChipYellowBackgroundBrush")
      $statusChipTextBlock.Foreground = $mainWindow.FindResource("ChipYellowForegroundBrush")
    }
    "Danger" {
      $statusChipTextBlock.Text = "エスカレ要"
      $statusChipContainer.Background = $mainWindow.FindResource("ChipRedBackgroundBrush")
      $statusChipTextBlock.Foreground = $mainWindow.FindResource("ChipRedForegroundBrush")
    }
  }
}

function Update-StatusChipByElapsedTime {
  if (-not $stopwatch.IsRunning) { Set-StatusChipState -state "Stopped"; return }

  $elapsedSeconds = [int][Math]::Floor($stopwatch.Elapsed.TotalSeconds)
  if ($elapsedSeconds -ge 15) { Set-StatusChipState -state "Danger"; return }
  if ($elapsedSeconds -ge 10) { Set-StatusChipState -state "Warning"; return }
  Set-StatusChipState -state "Working"
}

function Update-StopwatchButtons {
  if ($stopwatch.IsRunning) { $startStopButton.Content = "停止" } else { $startStopButton.Content = "開始" }
}

function Reset-StopwatchToStoppedState {
  try { $stopwatchTimer.Stop() } catch {}
  try { $stopwatch.Reset() } catch {}

  $stopwatchTimeTextBlock.Text = (Format-ElapsedTime -elapsedTime $stopwatch.Elapsed)
  Set-StatusChipState -state "Stopped"
  Update-StopwatchButtons
}

$stopwatchTimer.Add_Tick({
    $stopwatchTimeTextBlock.Text = (Format-ElapsedTime -elapsedTime $stopwatch.Elapsed)
    Update-StatusChipByElapsedTime
  }) | Out-Null

$startStopButton.Add_Click({
    if ($stopwatch.IsRunning) {
      $stopwatch.Stop()
      $stopwatchTimer.Stop()
      Set-StatusChipState -state "Stopped"
    }
    else {
      $stopwatch.Start()
      $stopwatchTimer.Start()
      Update-StatusChipByElapsedTime
    }

    $stopwatchTimeTextBlock.Text = (Format-ElapsedTime -elapsedTime $stopwatch.Elapsed)
    Update-StopwatchButtons
  }) | Out-Null

$resetButton.Add_Click({ Reset-StopwatchToStoppedState }) | Out-Null

$mainTabControl.Add_SelectionChanged({
    if ($script:IsBulkUpdateInProgress) { return }

    $selectedHeader = ""
    try { $selectedHeader = [string]$mainTabControl.SelectedItem.Header } catch {}

    $targetControlName = $null
    switch ($selectedHeader) {
      "Tab1" { $targetControlName = "FirstTabDropdown1" }
      "Tab2" { $targetControlName = "SecondTabDropdown1Special" }
      "Tab3" { $targetControlName = $null }
    }

    if ($targetControlName) {
      $mainWindow.Dispatcher.BeginInvoke([Action] { try { ($mainWindow.FindName($targetControlName)).Focus() | Out-Null } catch {} }) | Out-Null
    }
  }) | Out-Null

foreach ($definition in $formFieldDefinitions) {
  $control = $mainWindow.FindName($definition.ControlName)
  if ($definition.ControlType -eq "ComboBox") {
    $control.Add_SelectionChanged({ Sync-StateFromUserInterface -window $mainWindow }) | Out-Null
  }
  else {
    $control.Add_TextChanged({ Sync-StateFromUserInterface -window $mainWindow }) | Out-Null
  }
}

foreach ($definition in $formFieldDefinitions) {
  Attach-EnterKeyToMoveNext -control ($mainWindow.FindName($definition.ControlName))
}

$secondTabSpecialDropdown = [System.Windows.Controls.ComboBox]$mainWindow.FindName("SecondTabDropdown1Special")
$secondTabSpecialDropdown.Add_SelectionChanged({
    if ($script:IsBulkUpdateInProgress) { return }
    Update-SpecialAddressTextBoxVisibility -window $mainWindow
    Sync-StateFromUserInterface -window $mainWindow
  }) | Out-Null

$clearButton.Add_Click({
    try { [System.Windows.Clipboard]::Clear() } catch {}

    Clear-AllInputs -window $mainWindow
    Reset-StopwatchToStoppedState

    try { $mainTabControl.SelectedIndex = 0 } catch {}
    $mainWindow.Dispatcher.BeginInvoke([Action] { try { ($mainWindow.FindName("FirstTabDropdown1")).Focus() | Out-Null } catch {} }) | Out-Null
  }) | Out-Null

function Show-Drawer {
  $drawerOverlay.Visibility = "Visible"

  $openAnimation = New-Object System.Windows.Media.Animation.DoubleAnimation
  $openAnimation.From = -250
  $openAnimation.To = 0
  $openAnimation.Duration = [TimeSpan]::FromMilliseconds(180)
  $openAnimation.EasingFunction = New-Object System.Windows.Media.Animation.CubicEase
  $openAnimation.EasingFunction.EasingMode = [System.Windows.Media.Animation.EasingMode]::EaseOut

  $drawerTranslateTransform.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $openAnimation)
}

function Hide-Drawer {
  $closeAnimation = New-Object System.Windows.Media.Animation.DoubleAnimation
  $closeAnimation.From = $drawerTranslateTransform.X
  $closeAnimation.To = -250
  $closeAnimation.Duration = [TimeSpan]::FromMilliseconds(160)
  $closeAnimation.EasingFunction = New-Object System.Windows.Media.Animation.CubicEase
  $closeAnimation.EasingFunction.EasingMode = [System.Windows.Media.Animation.EasingMode]::EaseIn

  $closeAnimation.Add_Completed({ $drawerOverlay.Visibility = "Collapsed" }) | Out-Null
  $drawerTranslateTransform.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $closeAnimation)
}

function Show-Item1Screen {
  $headerTitleTextBlock.Text = "Item1"
  $placeholderScreen.Visibility = "Collapsed"
  $item1Screen.Visibility = "Visible"
}

function Show-PlaceholderScreen {
  param([Parameter(Mandatory)] [string] $titleText)

  $headerTitleTextBlock.Text = $titleText
  $placeholderTextBlock.Text = "$titleText`n準備中"
  $item1Screen.Visibility = "Collapsed"
  $placeholderScreen.Visibility = "Visible"
}

$hamburgerButton.Add_Click({ Show-Drawer }) | Out-Null
$drawerCloseButton.Add_Click({ Hide-Drawer }) | Out-Null

$drawerOverlay.Add_MouseDown({
    if ($_.OriginalSource -eq $drawerOverlay) {
      Hide-Drawer
      $_.Handled = $true
    }
  }) | Out-Null

$drawerPanel.Add_MouseDown({ $_.Handled = $true }) | Out-Null

$drawerItem1Button.Add_Click({ Show-Item1Screen; Hide-Drawer }) | Out-Null
$drawerItem2Button.Add_Click({ Show-PlaceholderScreen -titleText "Item2"; Hide-Drawer }) | Out-Null
$drawerItem3Button.Add_Click({ Show-PlaceholderScreen -titleText "Item3"; Hide-Drawer }) | Out-Null

Update-StopwatchButtons
$stopwatchTimeTextBlock.Text = (Format-ElapsedTime -elapsedTime $stopwatch.Elapsed)
Set-StatusChipState -state "Stopped"
Sync-UserInterfaceFromState -window $mainWindow

$mainWindow.Add_Loaded({
    try { Set-WindowBottomLeft -window $mainWindow } catch {}
    try { $drawerOverlay.Visibility = "Collapsed" } catch {}
    try { $drawerTranslateTransform.X = -250 } catch {}
    try { Show-Item1Screen } catch {}
  }) | Out-Null

$mainWindow.Add_Closing({
    try { Sync-StateFromUserInterface -window $mainWindow } catch {}
    try { $stopwatchTimer.Stop() } catch {}
    try { $stopwatch.Stop() } catch {}
  }) | Out-Null

$null = $mainWindow.ShowDialog()