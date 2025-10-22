# Create a new WPF window from an XML file and load all WPF elements and return them under one variable
function New-Window {
    param (
        [string]$filePath
    )

    try {
        [xml]$xaml = (Get-Content $filePath)
        $window = New-Object System.Collections.Hashtable
        $nodeReader = [System.Xml.XmlNodeReader]::New($xaml)
        $xamlReader = [Windows.Markup.XamlReader]::Load($nodeReader)
        [void]$window.Add('Window', $xamlReader)
        $elements = $xaml.SelectNodes("//*[@Name]")
        foreach ($element in $elements) {
            $varName = $element.Name
            $varValue = $window.Window.FindName($Element.Name)
            [void]$window.Add($varName, $varValue)
        }
        return $window
    } 
    catch {
        Show-ErrorMessageBox("Error building Xaml data or loading window data.`n$_")
        exit
    }
}

# Create new tabitem element
function New-Tab {
    param (
        [string]$name
    )

    $tabItem = New-Object System.Windows.Controls.TabItem
    $tabItem.Header = $name
    return $tabItem
}

# Create a new text label element
function New-Label {
    param (
        [string]$content,
        [string]$halign,
        [string]$valign
    )

    $label = New-Object System.Windows.Controls.Label
    $label.Content = $content
    $label.HorizontalAlignment = $halign
    $label.VerticalAlignment = $valign
    $label.Margin = New-Object System.Windows.Thickness(3)
    return $label
}

# Create a new tooltip element
function New-ToolTip {
    param (
        [string]$content
    )

    $tooltip = New-Object System.Windows.Controls.ToolTip
    $tooltip.Content = $content
    return $tooltip
}

# Create a new combo box element
function New-ComboBox {
    param (
        [string]$name,
        [System.String[]]$itemsSource,
        [string]$selectedItem
    )

    $comboBox = New-Object System.Windows.Controls.ComboBox
    $comboBox.Name = $name
    $comboBox.Margin = New-Object System.Windows.Thickness(5)
    $comboBox.ItemsSource = $itemsSource
    $comboBox.SelectedItem = $selectedItem
    return $comboBox
}

# Create a new text box element
function New-TextBox {
    param (
        [string]$name,
        [string]$text
    )

    $textBox = New-Object System.Windows.Controls.TextBox
    $textBox.Name = $name
    $textBox.Margin = New-Object System.Windows.Thickness(5)
    $textBox.Text = $text
    return $textBox
}

# Create a new check box element
function New-CheckBox {
    param (
        [string]$name,
        [bool]$isChecked
    )

    $checkbox = New-Object System.Windows.Controls.CheckBox
    $checkbox.Name = $name
    $checkbox.IsChecked = $isChecked
    return $checkbox
}

# Create a new button element
function New-Button {
    param (
        [string]$content,
        [string]$halign,
        [int]$width
    )
    
    $button = New-Object System.Windows.Controls.Button
    $button.Content = $content
    $button.Margin = New-Object System.Windows.Thickness(10)
    $button.HorizontalAlignment = $halign
    $button.Width = $width
    $button.IsDefault = $true
    return $button
}
