@echo off
cls

python -m nuitka ^
    main.UI.py ^
    --enable-plugin=upx ^
    --upx-binary=--best ^
    --standalone ^
    --windows-console-mode=disable ^
    --windows-icon-from-ico=assets/icon.ico ^
    --output-dir=../exe ^
    --windows-company-name=mememe2012 ^
    --windows-uac-admin ^
    --windows-product-name=Choice_UI ^
    --windows-file-version=1.3.2.0 ^
    --windows-product-version=1.3.2.0 ^
    --plugin-enable=tk-inter ^
    --jobs=8 ^
    --zig
    
pause