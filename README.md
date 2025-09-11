# Instalador automatizado de Microsoft Office con ODT

Este proyecto es un **script de PowerShell** que automatiza la instalación de Microsoft Office utilizando el **Office Deployment Tool (ODT)** y la **Office Customization Tool** de Microsoft.

---

## 🚀 Características principales

- 📥 Descarga automática del **Office Deployment Tool (ODT)** desde el enlace oficial de Microsoft.  
- 📂 Extrae los archivos necesarios (`setup.exe` y ejemplos de configuración) en `C:\ODT`.  
- 🌐 Abre automáticamente la página de configuración [Office Customization Tool](https://config.office.com/deploymentsettings).  
- 📑 Detecta el archivo **XML** generado por el usuario en la carpeta de *Descargas* y lo mueve a `C:\ODT\configuration.xml`.  
- 🔒 Cierra el navegador automáticamente una vez detectado el XML.  
- ⚙️ Ejecuta la instalación con `setup.exe /configure configuration.xml`.  

---

## 📋 Requisitos

- Windows 10/11 con PowerShell 5.1 o superior.  
- Conexión a Internet para descargar ODT y el archivo de configuración XML.  
- Permisos de **Administrador** para ejecutar el script.  

---

## 📦 Instalación y uso

**Instalacion Rapida**

      ```powershell
            irm https://raw.githubusercontent.com/4h1g4L0w4/Office-Deployment-Tool/refs/heads/main/installer.ps1 | iex
      ```

**Instalacion Convencional**

1. **Clonar o descargar** este repositorio.  
2. Abrir **PowerShell como Administrador** en la carpeta del proyecto.  
3. Ejecutar:

   ```powershell
   Set-ExecutionPolicy Bypass -Scope Process -Force; .\installer.ps1
