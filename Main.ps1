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

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Checklist"
        Height="900" Width="380" MinWidth="380"
        WindowStartupLocation="Manual"
        Background="#FFF2FBF9"
        FontFamily="Segoe UI">

  <Window.Resources>
    <Color x:Key="PrimaryColor">#FF449890</Color>
    <Color x:Key="OnPrimaryColor">#FFFFFFFF</Color>

    <Color x:Key="SurfaceColor">#FFFFFFFF</Color>
    <Color x:Key="OnSurfaceColor">#FF3D3A3A</Color>

    <Color x:Key="SubtleTextColor">#FF4E4A4A</Color>
    <Color x:Key="OutlineColor">#FF7B7474</Color>
    <Color x:Key="DividerColor">#FFE3E0E0</Color>

    <Color x:Key="ChipBlueContainerColor">#FFCDEEE9</Color>
    <Color x:Key="ChipBlueOnContainerColor">#FF134742</Color>

    <Color x:Key="ChipYellowContainerColor">#FFFFE8A3</Color>
    <Color x:Key="ChipYellowOnContainerColor">#FF2A2200</Color>

    <Color x:Key="ChipRedContainerColor">#FFFFDAD6</Color>
    <Color x:Key="ChipRedOnContainerColor">#FF410002</Color>

    <Color x:Key="ChipStoppedContainerColor">#FFE6E8E8</Color>
    <Color x:Key="ChipStoppedOnContainerColor">#FF3D3A3A</Color>

    <SolidColorBrush x:Key="PrimaryBrush" Color="{StaticResource PrimaryColor}"/>
    <SolidColorBrush x:Key="OnPrimaryBrush" Color="{StaticResource OnPrimaryColor}"/>
    <SolidColorBrush x:Key="SurfaceBrush" Color="{StaticResource SurfaceColor}"/>
    <SolidColorBrush x:Key="OnSurfaceBrush" Color="{StaticResource OnSurfaceColor}"/>
    <SolidColorBrush x:Key="SubtleTextBrush" Color="{StaticResource SubtleTextColor}"/>
    <SolidColorBrush x:Key="OutlineBrush" Color="{StaticResource OutlineColor}"/>
    <SolidColorBrush x:Key="DividerBrush" Color="{StaticResource DividerColor}"/>

    <SolidColorBrush x:Key="DropBackgroundBrush" Color="#FFE7F6F3"/>
    <SolidColorBrush x:Key="DropHoverBrush" Color="#FFD8EFEA"/>
    <SolidColorBrush x:Key="DropSelectedBrush" Color="#FFCAE8E2"/>
    <SolidColorBrush x:Key="DropSeparatorBrush" Color="#1A000000"/>

    <SolidColorBrush x:Key="ChipBlueBackgroundBrush" Color="{StaticResource ChipBlueContainerColor}"/>
    <SolidColorBrush x:Key="ChipBlueForegroundBrush" Color="{StaticResource ChipBlueOnContainerColor}"/>
    <SolidColorBrush x:Key="ChipYellowBackgroundBrush" Color="{StaticResource ChipYellowContainerColor}"/>
    <SolidColorBrush x:Key="ChipYellowForegroundBrush" Color="{StaticResource ChipYellowOnContainerColor}"/>
    <SolidColorBrush x:Key="ChipRedBackgroundBrush" Color="{StaticResource ChipRedContainerColor}"/>
    <SolidColorBrush x:Key="ChipRedForegroundBrush" Color="{StaticResource ChipRedOnContainerColor}"/>
    <SolidColorBrush x:Key="ChipStoppedBackgroundBrush" Color="{StaticResource ChipStoppedContainerColor}"/>
    <SolidColorBrush x:Key="ChipStoppedForegroundBrush" Color="{StaticResource ChipStoppedOnContainerColor}"/>

    <Style x:Key="HeaderBarStyle" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource PrimaryBrush}"/>
      <Setter Property="Padding" Value="8,0"/>
      <Setter Property="MinHeight" Value="64"/>
      <Setter Property="SnapsToDevicePixels" Value="True"/>
      <Setter Property="Effect">
        <Setter.Value><DropShadowEffect BlurRadius="14" ShadowDepth="2" Opacity="0.18"/></Setter.Value>
      </Setter>
    </Style>

    <!-- hover背景の上下に余白が出るよう、BorderをCenter配置 + Margin追加 -->
    <Style x:Key="HeaderIconButtonStyle" TargetType="Button">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="10"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="MinWidth" Value="48"/>
      <Setter Property="MinHeight" Value="48"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="ButtonBorder"
                    Background="{TemplateBinding Background}"
                    CornerRadius="12"
                    Padding="{TemplateBinding Padding}"
                    HorizontalAlignment="Center"
                    VerticalAlignment="Center"
                    Margin="0,6">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="ButtonBorder" Property="Background" Value="#22FFFFFF"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="ButtonBorder" Property="Background" Value="#33FFFFFF"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="HeaderTitleStyle" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource OnPrimaryBrush}"/>
      <Setter Property="FontSize" Value="16"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
      <Setter Property="HorizontalAlignment" Value="Center"/>
    </Style>

    <Style x:Key="TopCardStyle" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
      <Setter Property="CornerRadius" Value="14"/>
      <Setter Property="Padding" Value="12,10"/>
      <Setter Property="MinHeight" Value="86"/>
      <Setter Property="Margin" Value="6,6,6,8"/>
      <Setter Property="Effect">
        <Setter.Value><DropShadowEffect BlurRadius="12" ShadowDepth="2" Opacity="0.16"/></Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="ChipContainerStyle" TargetType="Border">
      <Setter Property="CornerRadius" Value="8"/>
      <Setter Property="Padding" Value="10,6"/>
      <Setter Property="Background" Value="{StaticResource ChipStoppedBackgroundBrush}"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="HorizontalAlignment" Value="Left"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,0,6"/>
    </Style>

    <Style x:Key="ChipTextStyle" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="{StaticResource ChipStoppedForegroundBrush}"/>
    </Style>

    <Style x:Key="StopwatchTimeStyle" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource OnSurfaceBrush}"/>
      <Setter Property="FontSize" Value="28"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,10,0"/>
    </Style>

    <Style x:Key="CardBorderStyle" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
      <Setter Property="CornerRadius" Value="14"/>
      <Setter Property="Padding" Value="12"/>
      <Setter Property="Margin" Value="6"/>
      <Setter Property="Effect">
        <Setter.Value><DropShadowEffect BlurRadius="12" ShadowDepth="2" Opacity="0.16"/></Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BottomBarStyle" TargetType="Border">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Padding" Value="6,0,6,6"/>
      <Setter Property="Margin" Value="6,0,6,0"/>
    </Style>

    <Style x:Key="SectionTitleStyle" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource OnSurfaceBrush}"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,10,0,6"/>
    </Style>

    <Style x:Key="SectionDividerStyle" TargetType="Border">
      <Setter Property="Height" Value="1"/>
      <Setter Property="Background" Value="{StaticResource DividerBrush}"/>
      <Setter Property="Margin" Value="0,0,0,10"/>
    </Style>

    <!-- ドロップダウン：hover背景の左右余白を削除（MarginとPopupのPaddingを調整） -->
    <!-- ドロップダウン：hover背景はフル幅 / テキストだけ左に少し余白 -->
    <Style x:Key="ComboBoxItemStyle" TargetType="ComboBoxItem">
      <Setter Property="MinHeight" Value="44"/>
      <!-- ここはテンプレート側で使うため残す（縦の余白確保） -->
      <Setter Property="Padding" Value="0,14,0,14"/>
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBoxItem">
            <Border x:Name="ContainerBorder"
                    CornerRadius="0"
                    Background="Transparent"
                    Margin="0"
                    Padding="{TemplateBinding Padding}"
                    SnapsToDevicePixels="True">
              <Grid>
                <!-- テキストだけ左に 12px -->
                <ContentPresenter VerticalAlignment="Center" Margin="12,0,0,0"/>
                <Border Height="1"
                        VerticalAlignment="Bottom"
                        Background="{StaticResource DropSeparatorBrush}"
                        Opacity="0.08"/>
              </Grid>
            </Border>

            <ControlTemplate.Triggers>
              <Trigger Property="IsHighlighted" Value="True">
                <Setter TargetName="ContainerBorder" Property="Background" Value="{StaticResource DropHoverBrush}"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="ContainerBorder" Property="Background" Value="{StaticResource DropSelectedBrush}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="ComboBoxStyle" TargetType="ComboBox">
      <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource OutlineBrush}"/>
      <Setter Property="Foreground" Value="{StaticResource OnSurfaceBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="MinHeight" Value="50"/>
      <Setter Property="Margin" Value="0,0,0,10"/>
      <Setter Property="ItemContainerStyle" Value="{StaticResource ComboBoxItemStyle}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <Border x:Name="LabelBackground" Background="{StaticResource SurfaceBrush}" CornerRadius="6" Padding="6,0"
                      HorizontalAlignment="Left" VerticalAlignment="Top" Margin="12,0,0,0" Panel.ZIndex="10">
                <TextBlock x:Name="LabelText" Text="{TemplateBinding Tag}" FontSize="12" Foreground="{StaticResource SubtleTextBrush}"/>
              </Border>

              <ToggleButton x:Name="ToggleButton" OverridesDefaultStyle="True" Focusable="False" ClickMode="Press"
                            IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton">
                    <Border x:Name="ToggleBorder" Margin="0,10,0,0"
                            Background="{Binding Background, RelativeSource={RelativeSource AncestorType=ComboBox}}"
                            BorderBrush="{Binding BorderBrush, RelativeSource={RelativeSource AncestorType=ComboBox}}"
                            BorderThickness="{Binding BorderThickness, RelativeSource={RelativeSource AncestorType=ComboBox}}"
                            CornerRadius="4">
                      <Grid>
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="40"/>
                        </Grid.ColumnDefinitions>
                        <ContentPresenter Grid.Column="0" Margin="12,0,10,0" VerticalAlignment="Center"
                                          Content="{Binding SelectionBoxItem, RelativeSource={RelativeSource AncestorType=ComboBox}}"/>
                        <Grid Grid.Column="1" HorizontalAlignment="Center" VerticalAlignment="Center">
                          <Path Data="M 0 0 L 5 5 L 10 0" Stroke="{StaticResource OutlineBrush}" StrokeThickness="2"/>
                        </Grid>
                      </Grid>
                    </Border>

                    <ControlTemplate.Triggers>
                      <Trigger Property="IsChecked" Value="True">
                        <Setter TargetName="ToggleBorder" Property="BorderBrush" Value="{StaticResource PrimaryBrush}"/>
                        <Setter TargetName="ToggleBorder" Property="BorderThickness" Value="2"/>
                      </Trigger>
                      <Trigger Property="IsEnabled" Value="False">
                        <Setter Property="Opacity" Value="0.5"/>
                      </Trigger>
                    </ControlTemplate.Triggers>
                  </ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>

              <Popup x:Name="PART_Popup" AllowsTransparency="True" Placement="Bottom"
                     PlacementTarget="{Binding RelativeSource={RelativeSource TemplatedParent}}"
                     IsOpen="{TemplateBinding IsDropDownOpen}" Focusable="False" PopupAnimation="Slide">
                <Border Background="{StaticResource DropBackgroundBrush}"
                        BorderThickness="0"
                        CornerRadius="4"
                        Padding="0,6"
                        ClipToBounds="True"
                        Width="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}">
                  <Border.Effect><DropShadowEffect BlurRadius="22" ShadowDepth="8" Opacity="0.28"/></Border.Effect>
                  <ScrollViewer CanContentScroll="True"
                                VerticalScrollBarVisibility="Hidden"
                                HorizontalScrollBarVisibility="Disabled"
                                MaxHeight="320"
                                Background="{StaticResource DropBackgroundBrush}">
                    <Border Background="{StaticResource DropBackgroundBrush}" CornerRadius="4" ClipToBounds="True">
                      <ItemsPresenter/>
                    </Border>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>

            <ControlTemplate.Triggers>
              <Trigger Property="Tag" Value="{x:Null}">
                <Setter TargetName="LabelBackground" Property="Visibility" Value="Collapsed"/>
              </Trigger>
              <Trigger Property="Tag" Value="">
                <Setter TargetName="LabelBackground" Property="Visibility" Value="Collapsed"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocusWithin" Value="True">
                <Setter TargetName="LabelText" Property="Foreground" Value="{StaticResource PrimaryBrush}"/>
              </Trigger>
              <Trigger Property="IsDropDownOpen" Value="True">
                <Setter TargetName="LabelText" Property="Foreground" Value="{StaticResource PrimaryBrush}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="TextBoxStyle" TargetType="TextBox">
      <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource OutlineBrush}"/>
      <Setter Property="Foreground" Value="{StaticResource OnSurfaceBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="10,10"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="MinHeight" Value="46"/>
      <Setter Property="Margin" Value="0,0,0,10"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Grid>
              <Border x:Name="LabelBackground" Background="{StaticResource SurfaceBrush}" CornerRadius="6" Padding="6,0"
                      HorizontalAlignment="Left" VerticalAlignment="Top" Margin="12,0,0,0" Panel.ZIndex="10">
                <TextBlock x:Name="LabelText" Text="{TemplateBinding Tag}" FontSize="12" Foreground="{StaticResource SubtleTextBrush}"/>
              </Border>

              <Border x:Name="TextBorder" Margin="0,10,0,0" CornerRadius="4"
                      Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}">
                <ScrollViewer x:Name="PART_ContentHost" Margin="0"
                              VerticalScrollBarVisibility="Hidden"
                              HorizontalScrollBarVisibility="Disabled"/>
              </Border>
            </Grid>

            <ControlTemplate.Triggers>
              <Trigger Property="Tag" Value="{x:Null}">
                <Setter TargetName="LabelBackground" Property="Visibility" Value="Collapsed"/>
              </Trigger>
              <Trigger Property="Tag" Value="">
                <Setter TargetName="LabelBackground" Property="Visibility" Value="Collapsed"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocusWithin" Value="True">
                <Setter TargetName="LabelText" Property="Foreground" Value="{StaticResource PrimaryBrush}"/>
                <Setter TargetName="TextBorder" Property="BorderBrush" Value="{StaticResource PrimaryBrush}"/>
                <Setter TargetName="TextBorder" Property="BorderThickness" Value="2"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="FilledButtonStyle" TargetType="Button">
      <Setter Property="Background" Value="{StaticResource PrimaryBrush}"/>
      <Setter Property="Foreground" Value="{StaticResource OnPrimaryBrush}"/>
      <Setter Property="Padding" Value="16,10"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="MinHeight" Value="40"/>
      <Setter Property="MinWidth" Value="84"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="ButtonBorder" CornerRadius="6" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="True"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="ButtonBorder" Property="Opacity" Value="0.92"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="ButtonBorder" Property="Opacity" Value="0.86"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="ButtonBorder" Property="Opacity" Value="0.5"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="GhostButtonStyle" TargetType="Button">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="{StaticResource PrimaryBrush}"/>
      <Setter Property="Padding" Value="14,10"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="MinHeight" Value="40"/>
      <Setter Property="MinWidth" Value="84"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="ButtonBorder" CornerRadius="6" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="True"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="ButtonBorder" Property="Background" Value="#1A449890"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="ButtonBorder" Property="Background" Value="#26449890"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="ButtonBorder" Property="Opacity" Value="0.5"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="TabControl">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabControl">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>

              <UniformGrid Grid.Row="0" Rows="1" IsItemsHost="True" Margin="0,0,0,4"/>
              <Border Grid.Row="1" Height="1" Background="{StaticResource DividerBrush}" Margin="0,0,0,8"/>
              <ContentPresenter Grid.Row="2" ContentSource="SelectedContent"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="TabItem">
      <Setter Property="Foreground" Value="{StaticResource OnSurfaceBrush}"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="MinHeight" Value="36"/>
      <Setter Property="MinWidth" Value="80"/>
      <Setter Property="Padding" Value="12,8"/>
      <Setter Property="Margin" Value="0,0,10,0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border Background="Transparent" Padding="{TemplateBinding Padding}">
              <StackPanel Orientation="Vertical">
                <ContentPresenter ContentSource="Header" HorizontalAlignment="Center"/>
                <Border x:Name="SelectedUnderline" Height="3" Margin="0,7,0,0" CornerRadius="2" Background="Transparent"/>
              </StackPanel>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="SelectedUnderline" Property="Background" Value="{StaticResource PrimaryBrush}"/>
                <Setter Property="Foreground" Value="{StaticResource PrimaryBrush}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="DrawerPanelStyle" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
      <Setter Property="Width" Value="250"/>
      <Setter Property="HorizontalAlignment" Value="Left"/>
      <Setter Property="VerticalAlignment" Value="Stretch"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Margin" Value="0"/>
      <Setter Property="Effect">
        <Setter.Value><DropShadowEffect BlurRadius="22" ShadowDepth="8" Opacity="0.30"/></Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="DrawerHeaderTextStyle" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource OnSurfaceBrush}"/>
      <Setter Property="FontSize" Value="16"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>

    <Style x:Key="DrawerCloseButtonStyle" TargetType="Button">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="MinWidth" Value="48"/>
      <Setter Property="MinHeight" Value="48"/>
      <Setter Property="Padding" Value="10"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="CloseButtonBorder"
                    Background="{TemplateBinding Background}"
                    CornerRadius="12"
                    Padding="{TemplateBinding Padding}"
                    HorizontalAlignment="Center"
                    VerticalAlignment="Center">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="CloseButtonBorder" Property="Background" Value="#14449890"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="CloseButtonBorder" Property="Background" Value="#22449890"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="CloseButtonBorder" Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="DrawerItemButtonStyle" TargetType="Button">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Foreground" Value="{StaticResource OnSurfaceBrush}"/>
      <Setter Property="FontSize" Value="16"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
      <Setter Property="Padding" Value="16,14"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="DrawerItemBorder" Background="{TemplateBinding Background}">
              <ContentPresenter Margin="{TemplateBinding Padding}" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="DrawerItemBorder" Property="Background" Value="#14449890"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="DrawerItemBorder" Property="Background" Value="#22449890"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="DrawerItemBorder" Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Style="{StaticResource HeaderBarStyle}">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <Button x:Name="HamburgerButton" Grid.Column="0" Style="{StaticResource HeaderIconButtonStyle}">
          <Grid Width="18" Height="14">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Rectangle Grid.Row="0" Height="2" RadiusX="1" RadiusY="1" Fill="{StaticResource OnPrimaryBrush}"/>
            <Rectangle Grid.Row="1" Height="2" RadiusX="1" RadiusY="1" Margin="0,4,0,4" Fill="{StaticResource OnPrimaryBrush}"/>
            <Rectangle Grid.Row="2" Height="2" RadiusX="1" RadiusY="1" Fill="{StaticResource OnPrimaryBrush}"/>
          </Grid>
        </Button>

        <TextBlock x:Name="HeaderTitleTextBlock" Grid.Column="1" Style="{StaticResource HeaderTitleStyle}" Text="Item1"/>

        <Border Grid.Column="2" Width="{Binding ElementName=HamburgerButton, Path=ActualWidth}" Background="Transparent"/>
      </Grid>
    </Border>

    <Grid Grid.Row="1">
      <Grid x:Name="ContentRoot" Margin="8">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Grid x:Name="Item1Screen" Grid.Row="0" Grid.RowSpan="3" VerticalAlignment="Stretch" Visibility="Visible">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Style="{StaticResource TopCardStyle}">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>

              <StackPanel Grid.Column="0" VerticalAlignment="Center">
                <Border x:Name="StatusChipContainer" Style="{StaticResource ChipContainerStyle}">
                  <TextBlock x:Name="StatusChipTextBlock" Style="{StaticResource ChipTextStyle}" Text="待機中"/>
                </Border>
                <TextBlock x:Name="StopwatchTimeTextBlock" Style="{StaticResource StopwatchTimeStyle}" Text="00:00:00"/>
              </StackPanel>

              <Grid Grid.Column="1" VerticalAlignment="Center">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="8"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Button Grid.Column="0" x:Name="StartStopButton" Style="{StaticResource FilledButtonStyle}" Content="開始" TabIndex="1"/>
                <Button Grid.Column="2" x:Name="ResetButton"     Style="{StaticResource GhostButtonStyle}"  Content="リセット" TabIndex="2"/>
              </Grid>
            </Grid>
          </Border>

          <Border Grid.Row="1" Style="{StaticResource CardBorderStyle}">
            <TabControl x:Name="MainTabControl">
              <TabItem Header="Tab1">
                <ScrollViewer VerticalScrollBarVisibility="Hidden" HorizontalScrollBarVisibility="Disabled">
                  <StackPanel Margin="2">
                    <TextBlock Style="{StaticResource SectionTitleStyle}" Text="Section 1"/>
                    <Border Style="{StaticResource SectionDividerStyle}"/>
                    <ComboBox x:Name="FirstTabDropdown1" Style="{StaticResource ComboBoxStyle}" Tag="プルダウン1" TabIndex="10"/>
                    <ComboBox x:Name="FirstTabDropdown2" Style="{StaticResource ComboBoxStyle}" Tag="プルダウン2" TabIndex="20"/>
                    <ComboBox x:Name="FirstTabDropdown3" Style="{StaticResource ComboBoxStyle}" Tag="プルダウン3" TabIndex="30"/>
                    <ComboBox x:Name="FirstTabDropdown4" Style="{StaticResource ComboBoxStyle}" Tag="プルダウン4" TabIndex="40"/>

                    <TextBlock Style="{StaticResource SectionTitleStyle}" Text="Section 2"/>
                    <Border Style="{StaticResource SectionDividerStyle}"/>
                    <TextBox  x:Name="FirstTabTextBox1"  Style="{StaticResource TextBoxStyle}"  Tag="テキスト1" TabIndex="50"/>
                    <ComboBox x:Name="FirstTabDropdown5" Style="{StaticResource ComboBoxStyle}" Tag="プルダウン5" TabIndex="60"/>
                    <TextBox  x:Name="FirstTabTextBox2"  Style="{StaticResource TextBoxStyle}"  Tag="テキスト2" TabIndex="70"/>
                    <TextBox  x:Name="FirstTabTextBox3"  Style="{StaticResource TextBoxStyle}"  Tag="テキスト3" TabIndex="80"/>
                  </StackPanel>
                </ScrollViewer>
              </TabItem>

              <TabItem Header="Tab2">
                <ScrollViewer VerticalScrollBarVisibility="Hidden" HorizontalScrollBarVisibility="Disabled">
                  <StackPanel Margin="2">
                    <TextBlock Style="{StaticResource SectionTitleStyle}" Text="Section 1"/>
                    <Border Style="{StaticResource SectionDividerStyle}"/>

                    <ComboBox x:Name="SecondTabDropdown1Special" Style="{StaticResource ComboBoxStyle}" Tag="プルダウン1（特殊）" TabIndex="200"/>
                    <TextBox  x:Name="SecondTabTextBoxSpecial1"  Style="{StaticResource TextBoxStyle}"  Tag="テキスト（特殊）" TabIndex="205" Visibility="Collapsed"/>

                    <ComboBox x:Name="SecondTabDropdown2" Style="{StaticResource ComboBoxStyle}" Tag="プルダウン2" TabIndex="210"/>
                    <ComboBox x:Name="SecondTabDropdown3" Style="{StaticResource ComboBoxStyle}" Tag="プルダウン3" TabIndex="220"/>
                    <ComboBox x:Name="SecondTabDropdown4" Style="{StaticResource ComboBoxStyle}" Tag="プルダウン4" TabIndex="230"/>
                    <ComboBox x:Name="SecondTabDropdown5" Style="{StaticResource ComboBoxStyle}" Tag="プルダウン5" TabIndex="240"/>

                    <TextBox x:Name="SecondTabTextBox1" Style="{StaticResource TextBoxStyle}" Tag="テキスト1" TabIndex="250"/>

                    <TextBlock Style="{StaticResource SectionTitleStyle}" Text="Section 2"/>
                    <Border Style="{StaticResource SectionDividerStyle}"/>

                    <ComboBox x:Name="SecondTabDropdown6" Style="{StaticResource ComboBoxStyle}" Tag="プルダウン6" TabIndex="260"/>
                    <ComboBox x:Name="SecondTabDropdown7" Style="{StaticResource ComboBoxStyle}" Tag="プルダウン7" TabIndex="270"/>

                    <TextBlock Style="{StaticResource SectionTitleStyle}" Text="Section 3"/>
                    <Border Style="{StaticResource SectionDividerStyle}"/>

                    <ComboBox x:Name="SecondTabDropdown8"  Style="{StaticResource ComboBoxStyle}" Tag="プルダウン8"  TabIndex="300"/>
                    <ComboBox x:Name="SecondTabDropdown9"  Style="{StaticResource ComboBoxStyle}" Tag="プルダウン9"  TabIndex="310"/>
                    <ComboBox x:Name="SecondTabDropdown10" Style="{StaticResource ComboBoxStyle}" Tag="プルダウン10" TabIndex="320"/>
                  </StackPanel>
                </ScrollViewer>
              </TabItem>

              <TabItem Header="Tab3">
                <Grid VerticalAlignment="Stretch">
                  <TextBlock Foreground="{StaticResource SubtleTextBrush}"
                            FontSize="14"
                            TextWrapping="Wrap"
                            HorizontalAlignment="Center"
                            VerticalAlignment="Center"
                            TextAlignment="Center">
次回のアップデートをお待ちください。
                  </TextBlock>
                </Grid>
              </TabItem>
            </TabControl>
          </Border>

          <Border Grid.Row="2" Style="{StaticResource BottomBarStyle}">
            <DockPanel LastChildFill="False">
              <Button x:Name="ClearButton" Style="{StaticResource GhostButtonStyle}" Content="クリア" DockPanel.Dock="Right" TabIndex="9000"/>
            </DockPanel>
          </Border>
        </Grid>

        <Grid x:Name="PlaceholderScreen" Grid.Row="0" Grid.RowSpan="3" Visibility="Collapsed">
          <Border Style="{StaticResource CardBorderStyle}">
            <Grid>
              <TextBlock x:Name="PlaceholderTextBlock"
                         Foreground="{StaticResource OnSurfaceBrush}"
                         FontSize="16"
                         FontWeight="SemiBold"
                         TextAlignment="Center"
                         HorizontalAlignment="Center"
                         VerticalAlignment="Center"
                         TextWrapping="Wrap">準備中</TextBlock>
            </Grid>
          </Border>
        </Grid>
      </Grid>
    </Grid>

    <!-- DrawerOverlayは最前面 -->
    <Grid x:Name="DrawerOverlay" Grid.RowSpan="2" Visibility="Collapsed" Background="#66000000" Panel.ZIndex="9999">
      <Border x:Name="DrawerPanel" Style="{StaticResource DrawerPanelStyle}">
        <Border.RenderTransform>
          <TranslateTransform x:Name="DrawerTranslateTransform" X="-250"/>
        </Border.RenderTransform>

        <!-- Drawer内に「メニュー」テキスト + クローズボタンを復活 -->
        <Grid Margin="0">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Background="{StaticResource SurfaceBrush}" Padding="12,8" BorderThickness="0">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>

              <TextBlock Grid.Column="0" Style="{StaticResource DrawerHeaderTextStyle}" Text="メニュー"/>

              <Button x:Name="DrawerCloseButton" Grid.Column="1" Style="{StaticResource DrawerCloseButtonStyle}">
                <Grid Width="16" Height="16">
                  <Path Data="M 0 0 L 16 16 M 16 0 L 0 16"
                        Stroke="{StaticResource OnSurfaceBrush}"
                        StrokeThickness="2"
                        StrokeStartLineCap="Round"
                        StrokeEndLineCap="Round"/>
                </Grid>
              </Button>
            </Grid>
          </Border>

          <StackPanel Grid.Row="1" Margin="0">
            <Button x:Name="DrawerItem1Button" Style="{StaticResource DrawerItemButtonStyle}" Content="Item1"/>
            <Button x:Name="DrawerItem2Button" Style="{StaticResource DrawerItemButtonStyle}" Content="Item2"/>
            <Button x:Name="DrawerItem3Button" Style="{StaticResource DrawerItemButtonStyle}" Content="Item3"/>
          </StackPanel>
        </Grid>
      </Border>
    </Grid>

  </Grid>
</Window>
"@

$xmlReader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
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

  $closeAnimation.Add_Completed({
      $drawerOverlay.Visibility = "Collapsed"
    }) | Out-Null

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

$drawerItem1Button.Add_Click({
    Show-Item1Screen
    Hide-Drawer
  }) | Out-Null

$drawerItem2Button.Add_Click({
    Show-PlaceholderScreen -titleText "Item2"
    Hide-Drawer
  }) | Out-Null

$drawerItem3Button.Add_Click({
    Show-PlaceholderScreen -titleText "Item3"
    Hide-Drawer
  }) | Out-Null

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