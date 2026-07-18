# BIM Manager Hub 🏢📊

![Profesión](https://img.shields.io/badge/Profesi%C3%B3n-Arquitectos%20T%C3%A9cnicos-2e7d32?logo=micro%3Abit&logoColor=white&style=flat-square)
![Role](https://img.shields.io/badge/Role-BIM%20%26%20ConTech-007ACC?logo=bim360&style=flat-square)
![Location](https://img.shields.io/badge/Location-A%20Coru%C3%B1a%20%F0%9F%8C%8A-005B94?logo=lighthouse&logoColor=white&style=flat-square)
![Sector](https://img.shields.io/badge/Sector-ConTech%20%7C%20AECO-E65100?logo=construct3&style=flat-square)
![BIM](https://img.shields.io/badge/BIM-IFC%20%2F%20openBIM-009688?style=flat-square)
![Maker](https://img.shields.io/badge/Maker-Software-red?logo=makerbot&style=flat-square)
![Hardware](https://img.shields.io/badge/Hardware---grey?style=flat-square)
![Windows](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&style=flat-square)
![Language](https://img.shields.io/badge/Language-Python-3776AB?logo=python&logoColor=white&style=flat-square)
![Stars](https://img.shields.io/github/stars/jmcaamanog/bim-manager-hub?style=flat-square&color=yellow&logo=github)
![License](https://img.shields.io/github/license/jmcaamanog/bim-manager-hub?style=flat-square&color=green)

(Arquitecto Técnico_JMC) Plataforma de escritorio integral para la gestión y administración en entornos BIM. Desarrollada en Python, combina un robusto backend de bases de datos relacionales con analítica de datos y una interfaz gráfica moderna, pensada para centralizar el flujo de trabajo de un BIM Manager.

## 🚀 Características Principales

* **Control de Acceso y Roles:** Sistema de autenticación integrado con base de datos local (SQLite). Gestiona usuarios, contraseñas encriptadas (`hashlib`) y roles de acceso a la plataforma.
* **Analítica y Dashboards:** Procesamiento de datos de proyectos mediante `Pandas` y `NumPy`, visualizados en tiempo real con gráficos dinámicos integrados en la interfaz gracias a `Matplotlib`.
* **Trazabilidad y Seguridad:** Sistema de registro de eventos (`logging`) que documenta cada acción en la app, sumado a una función de copias de seguridad automatizadas de los datos críticos.
* **Importación / Exportación Profesional:** Conexión bidireccional con formatos estándar del sector. Genera reportes formateados en Excel (`openpyxl`) y lee estructuras en formato CSV.
* **Despliegue Optimizado:** Código refactorizado y preparado con rutas relativas (`sys._MEIPASS`) para garantizar una compilación perfecta y sin errores usando PyInstaller.

## 🛠️ Stack Tecnológico

* **Frontend:** CustomTkinter (Dark Mode por defecto, diseño responsive).
* **Backend:** SQLite3 (Persistencia de datos), Hashlib (Seguridad).
* **Data Science:** Pandas, NumPy, Matplotlib.
* **Interoperabilidad:** Openpyxl (Excel), CSV, Pillow (Gestión de imágenes).

## ⚙️ Instalación y Uso

1. Clona este repositorio en tu equipo:
   ```bash
   git clone [https://github.com/TU_USUARIO/bim-manager-hub.git](https://github.com/TU_USUARIO/bim-manager-hub.git)

2. Navega al directorio del proyecto:
   ```bash
   cd bim-manager-hub

3. Instala las múltiples dependencias necesarias utilizando tu gestor de paquetes o entorno virtual:
   ```bash
   pip install customtkinter Pillow openpyxl matplotlib pandas numpy

4. Ejecuta la aplicación (la base de datos se inicializará automáticamente en el primer arranque):
   ```bash
   python bim_manager_app.py

## 👨‍💻 Autor

Jose Manuel Caamaño González | Arquitecto Técnico & BIM Manager.
Digital Product Lead | ConTech & Digital Twin SaaS | BIM, Energy Modeling & Sustainability | Data Analytics (SQL, Power BI)

Hecho con código y café desde A Coruña. ☕

Jose Manuel Caamaño González | [LinkedIn](https://www.linkedin.com/in/jmcaamanog/)
