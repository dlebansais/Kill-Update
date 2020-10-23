@echo off
echo Deploying ...
git fetch . master:deployment
git push origin deployment
:end