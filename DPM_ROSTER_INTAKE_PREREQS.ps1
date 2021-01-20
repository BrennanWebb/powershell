Install-PackageProvider -Name NuGet -scope CurrentUser -Force;

[array]$module = "ImportExcel","SqlServer";
foreach ($item in $module){
    If ($item -in (Get-InstalledModule).name){
        write-host "module "$item" exist";
    }
    else {
        install-module $item -Scope CurrentUser -force;
        write-host "module "$item" installed";
    }
};