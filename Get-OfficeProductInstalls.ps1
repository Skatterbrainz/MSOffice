function Get-OfficeProductInstalls {
    [CmdletBinding()]
    param ()
    $baseKey64 = "HKLM:\SOFTWARE\Microsoft\Office\NN\PRODNAME\InstallRoot"
    $baseKey32 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\NN\PRODNAME\InstallRoot"
    $products  = ("Access","Excel","Lync", "Outlook","PowerPoint","Publisher","Word","Project","Visio")
    $versions  = ("2003=11", "2007=12", "2010=14", "2013=15", "2016=16")
    $versions | Foreach-Object {
        $pvn = $($_ -split '=')[0]
        $pvv = $($_ -split '=')[1]
        $key64 = $baseKey64 -replace "NN", "$pvv`.0"
        $key32 = $baseKey32 -replace "NN", "$pvv`.0"
        foreach ($pn in $products) {
            $regkey64 = $key64 -replace "PRODNAME", $pn
            $regkey32 = $key32 -replace "PRODNAME", $pn
            if (Test-Path $regkey64) { $found64 = $True } else { $found64 = $False }
            if (Test-Path $regkey32) { $found32 = $True } else { $found32 = $False }
            switch ($pn) {
                "Visio" {
                    if ($found64) {
                        # search for std or pro
                    }
                    elseif ($found32) {
                        # search for std or pro
                    }
                }
                "Project" {
                    if ($found64) {
                        # search for std or pro
                    }
                    elseif ($found32) {
                        # search for std or pro
                    }
                }
            }
            [pscustomobject]@{
                Computer  = $env:COMPUTERNAME
                Product   = $pn
                Version   = $pvn
                VerNum    = "$pvv`.0"
                #KeyPath   = $regkey
                x64Installed = $found64
                x86Installed = $found32
            }
        }
    }
}
