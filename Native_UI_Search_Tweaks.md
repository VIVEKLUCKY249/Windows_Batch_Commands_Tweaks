### Search in same directory and ignore sub-directories
```bat
*.[ext] -folder:"[CurrentFolderName]\*"
*.mp4 -folder:"Downloads\*"
```
### Search in same directory's sub-directories only and ignore that directory
```bat
*.[ext] -folder:"[CurrentFolderName]\"
*.mp4 -folder:"Downloads\"
```
