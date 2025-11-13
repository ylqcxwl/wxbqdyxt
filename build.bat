@echo off
pip install -r requirements.txt
pyinstaller --onefile --windowed --noconsole ^
  --add-data "example_template.btw;." ^
  --add-data "config_pure.json;." ^
  --name=OuterBoxPrinter_PureData ^
  main_pure.py

echo 打包完成！输出：dist\OuterBoxPrinter_PureData.exe
pause