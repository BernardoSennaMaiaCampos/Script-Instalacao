# Conversão para .exe com ícone

ps2exe -inputFile "C:\caminho\do\arquivo.ps1" -outputFile "C:\caminho\Instalador_MultiSoftware.exe" -requireAdmin -iconFile "C:\caminho\icone.ico"


# Conversão para .exe sem ícone

ps2exe -inputFile "C:\Users\Documents\InstaladorAutomatico-1\install-software.ps1" -outputFile "C:\Users\Documents\Instalador.exe" -requireAdmin


ps2exe -inputFile install-software.ps1 -outputFile Instalador.exe -requireAdmin


ps2exe `
    -inputFile "C:\Users\Documents\InstaladorAutomatico-1\Instalador-Multi-Software-FINAL.ps1" `
    -outputFile "C:\Users\Documents\Instalador.exe" `
    -requireAdmin `
    -title "Instalador Multi-Software" `
    -version "2.2.0.0"
```

3. **Coloque os arquivos na mesma pasta do Instalador.exe:**
```
C:\Users\Documents\
├── Instalador.exe  ← Executável
├── FusionSigner_Instalador.exe
├── AssinadorLivreComMobileID.msi
├── AssinadorLivreComMobileID_Transform.mst
└── PJEOfficePro__Windows__x64__installer.exe





ps2exe `
     -inputFile "C:\Users\Documents\InstaladorAutomatico-1\main.ps1" `
     -outputFile "C:\Users\Documents\Instalador.exe" `
     -requireAdmin `
     -title "Instalador Multi-Software" `
     -version "2.0.0.0"



     
     
