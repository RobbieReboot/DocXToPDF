#REMEMBER THESE PATHS MUST POINT TO THE NUSPEC FILE - IN TC THIS IS RELATIVE i.e. "DocXToPDF.Core\DocXToPDF.Core.nuspec"
Write-Host "NuSpec Version : $(([xml]$nuSpec = Get-Content -Path 'DocXToPDF.Core.nuspec').package.metadata.version)"
Write-Output "##teamcity[setParameter name='env.nuspecVersion' value='$(([xml]$nuSpec = Get-Content -Path 'DocXToPDF.Core.nuspec').package.metadata.version)']"


