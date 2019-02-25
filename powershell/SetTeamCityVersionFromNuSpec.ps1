Write-Output "##teamcity[setParameter name='env.nuspecVersion' value='$(([xml]$nuSpec = Get-Content -Path 'DocXToPDF.Core.nuspec').package.metadata.version)']"


