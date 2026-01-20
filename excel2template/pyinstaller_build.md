PyInstallerによるWindows、macos(arm64)実行系の作成

## build時の開発環境

Windows、macos(arm64)ともに
- Python: 3.13.2
- pyinstaller: 6.17.0
- pyinstaller-hooks-contrib: 2025.11

## build

### Windows

```cmd
PS> pip install -r requirements.txt
PS> pip install pyinstaller
PS> pyinstaller --onefile --name excel2template.exe ./excel2template.py
```
buildした結果は./dist/excel2template.exe に出力されます。


### Mac OS(arm64)

```cmd
$ pip install -r requirements.txt
$ pip install pyinstaller
$ pyinstaller \
  --onefile \
  --name excel2template \
  ./excel2template.py
```

buildした結果は、./dist/excel2template に出力されます。

以上