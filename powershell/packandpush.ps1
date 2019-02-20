param (
    [string]$project = ".\DocXToPdf\DocXToPdf"
    )

#Build & pack
& "nuget.exe" pack -Build -properties Configuration=Release $project.csproj
& "nuget.exe" push *.nupkg -apikey woAsVXBI8v8W9UgIiNOr -Source http://3sq-ih-proget:8624/nuget/3SquaredCommonNuget/





