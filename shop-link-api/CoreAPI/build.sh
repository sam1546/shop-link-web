#!/bin/bash
echo Deleting Artifact folder if exist
rm -rf artifact

echo Creating output folder...
mkdir artifact

echo Publishing release....
dotnet publish --configuration Release --output artifact
#dotnet publish --self-contained=false -c Release --output artifact