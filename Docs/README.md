# 🛠️ Proyecto de Automatización Técnica – Instrumentación & Control

**Autor:** J. Livramento  
**Repositorio:** [ProjectAutomateSJL](https://github.com/javilivra/ProjectAutomateSJL)  
**Estado:** 🔧 *En construcción – documentación preliminar*

---

## 📌 Objetivo del Proyecto

Desarrollar macros y herramientas de automatización aplicadas a entornos técnicos de ingeniería, incluyendo:

- Extracción de atributos desde AutoCAD
- Manipulación inteligente de datos en Excel
- Automatizaciones con VBA orientadas a proyectos de Instrumentación y Control

---

## 📁 Estructura del repositorio

```bash
ProjectAutomateSJL/
├── Macros/           # Módulos .bas con código VBA versionado
├── excel/            # Plantillas de trabajo en Excel (próximamente)
├── autocad/          # Archivos DWG y recursos de prueba (próximamente)
├── Docs/             # Documentación técnica y flujos de trabajo
└── README.md         # Este archivo (manual del usuario - versión preliminar)
```
---

## 📖 Contenido del manual del usuario (en desarrollo)

Este README funcionará como manual oficial para el uso de las automatizaciones incluidas en el proyecto.

### 🗂️ Índice tentativo:

1. 🎯 Introducción y alcance del proyecto
2. ⚙️ Requisitos del entorno
   - Versiones compatibles de Excel / AutoCAD
   - Habilitación de macros y configuración de seguridad
3. 🧩 Estructura general del repositorio
4. 🚀 Instalación y carga de macros
   - Importar módulo `.bas` desde el editor VBA
   - Ubicación de archivos en la PC
5. 🧠 Descripción de funciones y rutinas
   - ExtraerAtributosBloqueInstrumentos()
   - CompletarSenalesYUnidades()
6. 🧪 Ejecución de la macro paso a paso
   - Qué hojas necesita
   - Qué archivos DWG debe tener cargados
   - Qué campos completa automáticamente
7. 🛠️ Resolución de errores comunes
8. 📦 Actualizaciones y control de versiones
9. 🧭 Buenas prácticas de uso en entornos de ingeniería
10. 📬 Contacto o contribución (si aplica en algún momento)

> ⚠️ Este índice se completará cuando se cierre la primera versión operativa y estable.


### ✅ Ejecución de pruebas

Para correr las pruebas automáticas ejecuta:

```bash
python -m unittest discover -s tests
```

