#requires -version 5.1
#requires -module GroupPolicy,ActiveDirectory

# Charger l'assembly WPF
Add-Type -AssemblyName PresentationFramework

# XAML pour l'interface utilisateur
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="GPO Manager" Height="450" Width="800">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Label Grid.Row="0" Content="Liste des GPOs" FontSize="16" HorizontalAlignment="Center"/>

        <ListView Grid.Row="1" Name="GpoListView" Margin="10">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Nom" DisplayMemberBinding="{Binding DisplayName}" Width="200"/>
                    <GridViewColumn Header="ID" DisplayMemberBinding="{Binding Id}" Width="300"/>
                    <GridViewColumn Header="OU Actuelle" DisplayMemberBinding="{Binding OU}" Width="300"/>
                </GridView>
            </ListView.View>
        </ListView>

        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center" Margin="10">
            <Button Content="Rafraîchir la liste" Name="RefreshButton" Margin="10" Padding="10"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Charger le XAML
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Récupérer les contrôles
$GpoListView = $Window.FindName("GpoListView")
$RefreshButton = $Window.FindName("RefreshButton")

# Fonction simplifiée pour récupérer les liens GPO
function Get-GPLink {
    [cmdletbinding()]
    [outputtype("PSObject")]
    Param(
        [parameter(Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$Name
    )

    # Récupérer les liens GPO pour les OUs uniquement
    $links = Get-GPInheritance -Target (Get-ADDomain).DistinguishedName -ErrorAction Stop | Select-Object -ExpandProperty GpoLinks
    $links += Get-ADOrganizationalUnit -Filter * | Get-GPInheritance -ErrorAction Stop | Select-Object -ExpandProperty GpoLinks

    # Filtrer par nom si spécifié
    if ($Name) {
        $links = $links | Where-Object { $_.DisplayName -like "*$Name*" }
    }

    # Retourner les résultats formatés
    $links | ForEach-Object {
        [PSCustomObject]@{
            DisplayName = $_.DisplayName
            GpoId       = $_.GpoId -replace ("{|}", "")
            Target      = $_.Target
            Enabled     = $_.Enabled
        }
    }
}

# Fonction pour charger les GPOs
function Load-GPOs {
    $GpoListView.Items.Clear()
    $GPOs = Get-GPO -All
    foreach ($GPO in $GPOs) {
        $GPLinks = Get-GPLink -Name $GPO.DisplayName
        $OU = ($GPLinks | Where-Object { $_.Target -match "^OU=" }).Target
        $GpoListView.Items.Add([PSCustomObject]@{
            DisplayName = $GPO.DisplayName
            Id = $GPO.Id
            OU = $OU
        })
    }
}

# Bouton pour rafraîchir la liste
$RefreshButton.Add_Click({
    Load-GPOs
    [System.Windows.MessageBox]::Show("Liste des GPOs rafraîchie!", "Succès")
})

# Charger les GPOs au démarrage
Load-GPOs

# Afficher la fenêtre
$Window.ShowDialog() | Out-Null