param (
    [string]$project = ".\DocXToPdf.Core\DocXToPdf.Core.csproj"
    )

#Build & pack
& "nuget.exe" pack -Build -properties Configuration=Release $project
#& "nuget.exe" push *.nupkg -apikey woAsVXBI8v8W9UgIiNOr -Source http://3sq-ih-proget:8624/nuget/3SquaredCommonNuget/

#nuget.exe pack -Build -properties Configuration=Release .\DocXToPdf.Core\DocXToPdf.Core.nusp