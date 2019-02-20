param (
    [string]$project = ".\docx2pdf"
    )

#Build & pack
& "nuget.exe" pack -build -properties Configuration=Release $project.nuspec
& "nuget.exe" push *.nupkg -apikey woAsVXBI8v8W9UgIiNOr -Source http://3sq-ih-proget:8624/nuget/3SquaredCommonNuget/





