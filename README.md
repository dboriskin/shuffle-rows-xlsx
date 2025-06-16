
# Shuffle Rows XLSX

## Как использовать

1️⃣ Залей в приватный репозиторий `shuffle-rows-xlsx` на GitHub.  
2️⃣ Перейди во вкладку **Actions**, запусти workflow `Build Installer`.  
3️⃣ После завершения забери готовый `Setup_ShuffleRowsXlsx.exe` в артефактах.  
4️⃣ Настя запускает установщик — выбирает папку, ярлык появляется на рабочем столе.  

## Что внутри
- `shuffle-rows-xslx.py` — финальный скрипт с автосозданием папок и шаблоном  
- `inno_setup.iss` — скрипт Inno Setup для создания установщика  
- `.github/workflows/build-installer.yml` — Actions workflow для сборки EXE + инсталлятора
