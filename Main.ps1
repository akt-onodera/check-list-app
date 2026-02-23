#requires -version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$script:bulk = $false

$def = @{
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

$script:st = @{}
foreach ($k in $def.Keys) { $script:st[$k] = $def[$k] }

$script:sel = @{ d1 = ""; d2 = ""; d3 = "" }

function Get-CbText {
  param([Parameter(Mandatory)][System.Windows.Controls.ComboBox]$cb)
  if ($cb.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) { return [string]$cb.SelectedItem.Content }
  ""
}

function Set-CbByText {
  param(
    [Parameter(Mandatory)][System.Windows.Controls.ComboBox]$cb,
    [AllowEmptyString()][string]$t = ""
  )
  if ([string]::IsNullOrEmpty($t)) { $cb.SelectedIndex = -1; $cb.SelectedItem = $null; return }
  foreach ($it in $cb.Items) {
    if ($it -is [System.Windows.Controls.ComboBoxItem] -and [string]$it.Content -eq $t) { $cb.SelectedItem = $it; return }
  }
  $cb.SelectedIndex = -1
  $cb.SelectedItem = $null
}

function Set-CbItems {
  param(
    [Parameter(Mandatory)][System.Windows.Controls.ComboBox]$cb,
    [Parameter(Mandatory)]$items,
    [AllowEmptyString()][string]$keep = ""
  )

  $arr = @()
  if ($items -is [string]) { $arr = @([string]$items) }
  elseif ($items -is [System.Collections.IEnumerable]) { foreach ($v in $items) { $arr += [string]$v } }
  else { $arr = @([string]$items) }

  $cb.Items.Clear()
  foreach ($s in $arr) {
    $cbi = New-Object System.Windows.Controls.ComboBoxItem
    $cbi.Content = $s
    $cb.Items.Add($cbi) | Out-Null
  }
  Set-CbByText -cb $cb -t $keep
}

function Focus-Next {
  param([Parameter(Mandatory)][System.Windows.DependencyObject]$from)
  $req = New-Object System.Windows.Input.TraversalRequest([System.Windows.Input.FocusNavigationDirection]::Next)
  $null = $from.MoveFocus($req)
}

function Attach-EnterNext {
  param([Parameter(Mandatory)]$c)

  if ($c -is [System.Windows.Controls.TextBox]) {
    $c.AcceptsReturn = $false
    $c.Add_PreviewKeyDown({
        if ($_.Key -eq [System.Windows.Input.Key]::Enter) { $_.Handled = $true; Focus-Next -from $this }
      }) | Out-Null
    return
  }

  if ($c -is [System.Windows.Controls.ComboBox]) {
    $c.Add_PreviewKeyDown({
        if ($_.Key -eq [System.Windows.Input.Key]::Enter -and -not $this.IsDropDownOpen) { $_.Handled = $true; Focus-Next -from $this }
      }) | Out-Null
  }
}

function Set-BottomLeft {
  param([Parameter(Mandatory)][System.Windows.Window]$w)
  try {
    $wa = [System.Windows.SystemParameters]::WorkArea
    $w.Left = $wa.Left + 10
    $w.Top = $wa.Bottom - $w.ActualHeight - 10
  }
  catch {}
}

function Fmt-Time {
  param([Parameter(Mandatory)][TimeSpan]$ts)
  "{0:00}:{1:00}:{2:00}" -f [int]$ts.TotalHours, $ts.Minutes, $ts.Seconds
}

$defs = @(
  @{ n = "FirstTabDropdown1"; k = "FirstTabDropdown1"; t = "cb" }
  @{ n = "FirstTabDropdown2"; k = "FirstTabDropdown2"; t = "cb" }
  @{ n = "FirstTabDropdown3"; k = "FirstTabDropdown3"; t = "cb" }
  @{ n = "FirstTabDropdown4"; k = "FirstTabDropdown4"; t = "cb" }
  @{ n = "FirstTabTextBox1"; k = "FirstTabTextBox1"; t = "tb" }
  @{ n = "FirstTabDropdown5"; k = "FirstTabDropdown5"; t = "cb" }
  @{ n = "FirstTabTextBox2"; k = "FirstTabTextBox2"; t = "tb" }
  @{ n = "FirstTabTextBox3"; k = "FirstTabTextBox3"; t = "tb" }

  @{ n = "SecondTabDropdown1Special"; k = "SecondTabDropdown1Special"; t = "cb" }
  @{ n = "SecondTabTextBoxSpecial1"; k = "SecondTabTextBoxSpecial1"; t = "tb" }

  @{ n = "SecondTabDropdown2"; k = "SecondTabDropdown2"; t = "cb" }
  @{ n = "SecondTabDropdown3"; k = "SecondTabDropdown3"; t = "cb" }
  @{ n = "SecondTabDropdown4"; k = "SecondTabDropdown4"; t = "cb" }
  @{ n = "SecondTabDropdown5"; k = "SecondTabDropdown5"; t = "cb" }
  @{ n = "SecondTabTextBox1"; k = "SecondTabTextBox1"; t = "tb" }
  @{ n = "SecondTabDropdown6"; k = "SecondTabDropdown6"; t = "cb" }
  @{ n = "SecondTabDropdown7"; k = "SecondTabDropdown7"; t = "cb" }
  @{ n = "SecondTabDropdown8"; k = "SecondTabDropdown8"; t = "cb" }
  @{ n = "SecondTabDropdown9"; k = "SecondTabDropdown9"; t = "cb" }
  @{ n = "SecondTabDropdown10"; k = "SecondTabDropdown10"; t = "cb" }
)

$opt = @{
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

$script:dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$xaml = Join-Path $script:dir "MainWindow.xaml"
$script:w = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml](Get-Content -Raw -Encoding UTF8 $xaml))))

function Upd-SpecialTb {
  param([Parameter(Mandatory)][System.Windows.Window]$w)
  $cb = [System.Windows.Controls.ComboBox]$w.FindName("SecondTabDropdown1Special")
  $tb = [System.Windows.Controls.TextBox]$w.FindName("SecondTabTextBoxSpecial1")
  $t = if ($cb) { Get-CbText -cb $cb } else { "" }

  if ($t -eq "選択肢C") {
    $tb.Visibility = "Visible"
    $w.Dispatcher.BeginInvoke([Action] { try { $tb.Focus() | Out-Null; $tb.SelectAll() } catch {} }) | Out-Null
  }
  else {
    $tb.Visibility = "Collapsed"
    if ($tb.Text -ne "") { $tb.Text = "" }
  }
}

function Apply-Opts {
  param([Parameter(Mandatory)][System.Windows.Window]$w)
  foreach ($n in $opt.Keys) {
    $cb = [System.Windows.Controls.ComboBox]$w.FindName($n)
    if ($cb) {
      $keep = Get-CbText -cb $cb
      Set-CbItems -cb $cb -items @($opt[$n]) -keep $keep
    }
  }
}

function Sync-Ui {
  param([Parameter(Mandatory)][System.Windows.Window]$w)
  $script:bulk = $true
  try {
    Apply-Opts -w $w
    foreach ($d in $defs) {
      $c = $w.FindName($d.n)
      $v = [string]$script:st[$d.k]
      if ($d.t -eq "cb") { Set-CbByText -cb $c -t $v } else { $c.Text = $v }
    }
    Upd-SpecialTb -w $w
  }
  finally { $script:bulk = $false }
}

function Sync-St {
  param([Parameter(Mandatory)][System.Windows.Window]$w)
  if ($script:bulk) { return }
  foreach ($d in $defs) {
    $c = $w.FindName($d.n)
    if ($d.t -eq "cb") { $script:st[$d.k] = Get-CbText -cb $c } else { $script:st[$d.k] = $c.Text }
  }
  if ($script:st["SecondTabDropdown1Special"] -ne "選択肢C") { $script:st["SecondTabTextBoxSpecial1"] = "" }
}

function Clear-All {
  param([Parameter(Mandatory)][System.Windows.Window]$w)
  $script:bulk = $true
  try {
    foreach ($k in $def.Keys) { $script:st[$k] = "" }
    foreach ($d in $defs) {
      $c = $w.FindName($d.n)
      if ($d.t -eq "cb") { $c.SelectedIndex = -1; $c.SelectedItem = $null } else { $c.Text = "" }
    }
    Upd-SpecialTb -w $w
  }
  finally { $script:bulk = $false }
}

function Set-ModalCbStyle {
  param([Parameter(Mandatory)][System.Windows.Controls.ComboBox]$cb, [Parameter(Mandatory)][bool]$ok)
  if ($ok) { $cb.Style = $script:w.FindResource("UnifiedComboBoxStyle") }
  else { $cb.Style = $script:w.FindResource("UnifiedComboBoxErrorStyle") }
}

function Hide-Modal {
  $ov = [System.Windows.Controls.Grid]$script:w.FindName("StartupModalOverlay")
  if ($ov) { $ov.Visibility = "Collapsed" }
}

$script:modalInit = $true
$script:modalTouched = $false
$script:modalFromUser = $false

function Modal-SetTouched([bool]$v) { $script:modalTouched = $v }

function Show-Modal {
  param([Parameter(Mandatory)][bool]$fromUser)

  $ov = [System.Windows.Controls.Grid]$script:w.FindName("StartupModalOverlay")
  if (-not $ov) { return }

  $script:modalFromUser = $fromUser
  Modal-SetTouched $false

  $cb1 = [System.Windows.Controls.ComboBox]$script:w.FindName("StartupDropdown1")
  $cb2 = [System.Windows.Controls.ComboBox]$script:w.FindName("StartupDropdown2")
  $cb3 = [System.Windows.Controls.ComboBox]$script:w.FindName("StartupDropdown3")
  $btnCancel = [System.Windows.Controls.Button]$script:w.FindName("StartupCancelButton")

  if ($btnCancel) {
    if ($fromUser) { $btnCancel.Visibility = "Visible" }
    else { $btnCancel.Visibility = "Collapsed" }
  }

  if ($cb1) { Set-CbByText -cb $cb1 -t $script:sel.d1; $cb1.Style = $script:w.FindResource("UnifiedComboBoxStyle") }
  if ($cb2) { Set-CbByText -cb $cb2 -t $script:sel.d2; $cb2.Style = $script:w.FindResource("UnifiedComboBoxStyle") }
  if ($cb3) { Set-CbByText -cb $cb3 -t $script:sel.d3; $cb3.Style = $script:w.FindResource("UnifiedComboBoxStyle") }

  $ov.Visibility = "Visible"
}

function Modal-Validate([bool]$applyStyle) {
  $cb1 = [System.Windows.Controls.ComboBox]$script:w.FindName("StartupDropdown1")
  $cb2 = [System.Windows.Controls.ComboBox]$script:w.FindName("StartupDropdown2")
  $cb3 = [System.Windows.Controls.ComboBox]$script:w.FindName("StartupDropdown3")

  $t1 = if ($cb1) { Get-CbText -cb $cb1 } else { "" }
  $t2 = if ($cb2) { Get-CbText -cb $cb2 } else { "" }
  $t3 = if ($cb3) { Get-CbText -cb $cb3 } else { "" }

  $ok1 = -not [string]::IsNullOrEmpty($t1)
  $ok2 = -not [string]::IsNullOrEmpty($t2)
  $ok3 = -not [string]::IsNullOrEmpty($t3)

  if ($applyStyle) {
    if ($cb1) { Set-ModalCbStyle -cb $cb1 -ok $ok1 }
    if ($cb2) { Set-ModalCbStyle -cb $cb2 -ok $ok2 }
    if ($cb3) { Set-ModalCbStyle -cb $cb3 -ok $ok3 }
  }

  return @{ ok = ($ok1 -and $ok2 -and $ok3); t1 = $t1; t2 = $t2; t3 = $t3 }
}

function Init-Modal {
  $cb1 = [System.Windows.Controls.ComboBox]$script:w.FindName("StartupDropdown1")
  $cb2 = [System.Windows.Controls.ComboBox]$script:w.FindName("StartupDropdown2")
  $cb3 = [System.Windows.Controls.ComboBox]$script:w.FindName("StartupDropdown3")
  $btnOk = [System.Windows.Controls.Button]$script:w.FindName("StartupConfirmButton")
  $btnNg = [System.Windows.Controls.Button]$script:w.FindName("StartupCancelButton")

  $opt2 = @("選択肢1", "選択肢2", "選択肢3", "選択肢4", "選択肢5", "選択肢6", "選択肢7", "選択肢8", "選択肢9", "選択肢10")

  if ($cb1) {
    Set-CbItems -cb $cb1 -items @("選択肢A", "選択肢B") -keep ""
    $cb1.Add_SelectionChanged({
        if (-not $script:modalTouched) { return }
        $t = Get-CbText -cb $this
        Set-ModalCbStyle -cb $this -ok:(-not [string]::IsNullOrEmpty($t))
      }) | Out-Null
  }

  if ($cb2) {
    Set-CbItems -cb $cb2 -items $opt2 -keep ""
    $cb2.Add_SelectionChanged({
        if (-not $script:modalTouched) { return }
        $t = Get-CbText -cb $this
        Set-ModalCbStyle -cb $this -ok:(-not [string]::IsNullOrEmpty($t))
      }) | Out-Null
  }

  if ($cb3) {
    Set-CbItems -cb $cb3 -items @("選択肢A", "選択肢B") -keep ""
    $cb3.Add_SelectionChanged({
        if (-not $script:modalTouched) { return }
        $t = Get-CbText -cb $this
        Set-ModalCbStyle -cb $this -ok:(-not [string]::IsNullOrEmpty($t))
      }) | Out-Null
  }

  if ($btnOk) {
    $btnOk.Add_Click({
        Modal-SetTouched $true
        $r = Modal-Validate -applyStyle:$true
        if (-not $r.ok) { return }
        $script:sel.d1 = $r.t1
        $script:sel.d2 = $r.t2
        $script:sel.d3 = $r.t3
        Hide-Modal
      }) | Out-Null
  }

  if ($btnNg) {
    $btnNg.Add_Click({
        Modal-SetTouched $true
        $r = Modal-Validate -applyStyle:$true
        if (-not $r.ok) { return }
        Hide-Modal
      }) | Out-Null
  }
}

$sw = New-Object System.Diagnostics.Stopwatch
$tm = New-Object System.Windows.Threading.DispatcherTimer
$tm.Interval = [TimeSpan]::FromMilliseconds(200)

$tSw = [System.Windows.Controls.TextBlock]$script:w.FindName("StopwatchTimeTextBlock")
$bSw = [System.Windows.Controls.Button]$script:w.FindName("StartStopButton")
$bRs = [System.Windows.Controls.Button]$script:w.FindName("ResetButton")
$bCl = [System.Windows.Controls.Button]$script:w.FindName("ClearButton")

$chipBd = [System.Windows.Controls.Border]$script:w.FindName("StatusChipContainer")
$chipTx = [System.Windows.Controls.TextBlock]$script:w.FindName("StatusChipTextBlock")
$tabs = [System.Windows.Controls.TabControl]$script:w.FindName("MainTabControl")

$btnHam = [System.Windows.Controls.Button]$script:w.FindName("HamburgerButton")
$btnUsr = [System.Windows.Controls.Button]$script:w.FindName("UserButton")
$imgUsr = [System.Windows.Controls.Image]$script:w.FindName("UserIconImage")

$ovDr = [System.Windows.Controls.Grid]$script:w.FindName("DrawerOverlay")
$pnDr = [System.Windows.Controls.Border]$script:w.FindName("DrawerPanel")
$trDr = [System.Windows.Media.TranslateTransform]$script:w.FindName("DrawerTranslateTransform")

$bDrX = [System.Windows.Controls.Button]$script:w.FindName("DrawerCloseButton")
$bD1 = [System.Windows.Controls.Button]$script:w.FindName("DrawerItem1Button")
$bD2 = [System.Windows.Controls.Button]$script:w.FindName("DrawerItem2Button")
$bD3 = [System.Windows.Controls.Button]$script:w.FindName("DrawerItem3Button")

$ttl = [System.Windows.Controls.TextBlock]$script:w.FindName("HeaderTitleTextBlock")
$scr1 = [System.Windows.Controls.Grid]$script:w.FindName("Item1Screen")
$scrP = [System.Windows.Controls.Grid]$script:w.FindName("PlaceholderScreen")
$txtP = [System.Windows.Controls.TextBlock]$script:w.FindName("PlaceholderTextBlock")

function Set-Chip([ValidateSet("Stopped", "Working", "Warning", "Danger")][string]$s) {
  switch ($s) {
    "Stopped" { $chipTx.Text = "待機中"; $chipBd.Background = $script:w.FindResource("ChipStoppedBackgroundBrush"); $chipTx.Foreground = $script:w.FindResource("ChipStoppedForegroundBrush") }
    "Working" { $chipTx.Text = "対応中"; $chipBd.Background = $script:w.FindResource("ChipBlueBackgroundBrush"); $chipTx.Foreground = $script:w.FindResource("ChipBlueForegroundBrush") }
    "Warning" { $chipTx.Text = "保留延伸"; $chipBd.Background = $script:w.FindResource("ChipYellowBackgroundBrush"); $chipTx.Foreground = $script:w.FindResource("ChipYellowForegroundBrush") }
    "Danger" { $chipTx.Text = "エスカレ要"; $chipBd.Background = $script:w.FindResource("ChipRedBackgroundBrush"); $chipTx.Foreground = $script:w.FindResource("ChipRedForegroundBrush") }
  }
}

function Upd-Chip {
  if (-not $sw.IsRunning) { Set-Chip "Stopped"; return }
  $sec = [int][Math]::Floor($sw.Elapsed.TotalSeconds)
  if ($sec -ge 15) { Set-Chip "Danger"; return }
  if ($sec -ge 10) { Set-Chip "Warning"; return }
  Set-Chip "Working"
}

function Upd-SwBtn { if ($sw.IsRunning) { $bSw.Content = "停止" } else { $bSw.Content = "開始" } }

function Reset-Sw {
  $tm.Stop()
  $sw.Reset()
  $tSw.Text = (Fmt-Time -ts $sw.Elapsed)
  Set-Chip "Stopped"
  Upd-SwBtn
}

$tm.Add_Tick({
    $tSw.Text = (Fmt-Time -ts $sw.Elapsed)
    Upd-Chip
  }) | Out-Null

$bSw.Add_Click({
    if ($sw.IsRunning) { $sw.Stop(); $tm.Stop(); Set-Chip "Stopped" }
    else { $sw.Start(); $tm.Start(); Upd-Chip }
    $tSw.Text = (Fmt-Time -ts $sw.Elapsed)
    Upd-SwBtn
  }) | Out-Null

$bRs.Add_Click({ Reset-Sw }) | Out-Null

$tabs.Add_SelectionChanged({
    if ($script:bulk) { return }
    $h = ""
    try { $h = [string]$tabs.SelectedItem.Header } catch {}
    $t = $null
    switch ($h) {
      "Tab1" { $t = "FirstTabDropdown1" }
      "Tab2" { $t = "SecondTabDropdown1Special" }
      "Tab3" { $t = $null }
    }
    if ($t) { $script:w.Dispatcher.BeginInvoke([Action] { try { ($script:w.FindName($t)).Focus() | Out-Null } catch {} }) | Out-Null }
  }) | Out-Null

foreach ($d in $defs) {
  $c = $script:w.FindName($d.n)
  if ($d.t -eq "cb") { $c.Add_SelectionChanged({ Sync-St -w $script:w }) | Out-Null }
  else { $c.Add_TextChanged({ Sync-St -w $script:w }) | Out-Null }
  Attach-EnterNext -c $c
}

$cbSpec = [System.Windows.Controls.ComboBox]$script:w.FindName("SecondTabDropdown1Special")
$cbSpec.Add_SelectionChanged({
    if ($script:bulk) { return }
    Upd-SpecialTb -w $script:w
    Sync-St -w $script:w
  }) | Out-Null

$bCl.Add_Click({
    [System.Windows.Clipboard]::Clear()
    Clear-All -w $script:w
    Reset-Sw
    try { $tabs.SelectedIndex = 0 } catch {}
    $script:w.Dispatcher.BeginInvoke([Action] { try { ($script:w.FindName("FirstTabDropdown1")).Focus() | Out-Null } catch {} }) | Out-Null
  }) | Out-Null

function Show-Drawer {
  $ovDr.Visibility = "Visible"
  $a = New-Object System.Windows.Media.Animation.DoubleAnimation
  $a.From = -250; $a.To = 0; $a.Duration = [TimeSpan]::FromMilliseconds(180)
  $a.EasingFunction = New-Object System.Windows.Media.Animation.CubicEase
  $a.EasingFunction.EasingMode = [System.Windows.Media.Animation.EasingMode]::EaseOut
  $trDr.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $a)
}

function Hide-Drawer {
  $a = New-Object System.Windows.Media.Animation.DoubleAnimation
  $a.From = $trDr.X; $a.To = -250; $a.Duration = [TimeSpan]::FromMilliseconds(160)
  $a.EasingFunction = New-Object System.Windows.Media.Animation.CubicEase
  $a.EasingFunction.EasingMode = [System.Windows.Media.Animation.EasingMode]::EaseIn
  $a.Add_Completed({ $ovDr.Visibility = "Collapsed" }) | Out-Null
  $trDr.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $a)
}

function Show-Item1 { $ttl.Text = "Item1"; $scrP.Visibility = "Collapsed"; $scr1.Visibility = "Visible" }
function Show-Ph([Parameter(Mandatory)][string]$t) { $ttl.Text = $t; $txtP.Text = "$t`n準備中"; $scr1.Visibility = "Collapsed"; $scrP.Visibility = "Visible" }

$btnHam.Add_Click({ Show-Drawer }) | Out-Null
$bDrX.Add_Click({ Hide-Drawer }) | Out-Null

$ovDr.Add_MouseDown({
    if ($_.OriginalSource -eq $ovDr) { Hide-Drawer; $_.Handled = $true }
  }) | Out-Null

$pnDr.Add_MouseDown({ $_.Handled = $true }) | Out-Null

$bD1.Add_Click({ Show-Item1; Hide-Drawer }) | Out-Null
$bD2.Add_Click({ Show-Ph -t "Item2"; Hide-Drawer }) | Out-Null
$bD3.Add_Click({ Show-Ph -t "Item3"; Hide-Drawer }) | Out-Null

function Set-UsrIcon {
  if (-not $imgUsr) { return }
  $p = Join-Path $script:dir "images\account_circle.png"
  if (-not (Test-Path $p)) { return }
  $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
  $bmp.BeginInit()
  $bmp.UriSource = New-Object System.Uri($p, [System.UriKind]::Absolute)
  $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
  $bmp.EndInit()
  $bmp.Freeze()
  $imgUsr.Source = $bmp
}

if ($btnUsr) { $btnUsr.Add_Click({ Show-Modal -fromUser:$true }) | Out-Null }

Init-Modal
Apply-Opts -w $script:w
Sync-Ui -w $script:w

Upd-SwBtn
$tSw.Text = (Fmt-Time -ts $sw.Elapsed)
Set-Chip "Stopped"

$script:w.Add_Loaded({
    Set-BottomLeft -w $script:w
    $ovDr.Visibility = "Collapsed"
    $trDr.X = -250
    Show-Item1
    Set-UsrIcon
    Show-Modal -fromUser:$false
  }) | Out-Null

$script:w.Add_Closing({
    try { Sync-St -w $script:w } catch {}
    $tm.Stop()
    $sw.Stop()
  }) | Out-Null

$null = $script:w.ShowDialog()