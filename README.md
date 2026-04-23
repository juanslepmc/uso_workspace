
# 📊 Uso Workspace – Procesamiento de Usuarios

Este proyecto contiene un conjunto de scripts en Python orientados al **procesamiento, separación y cruce de información de usuarios**, a partir de distintos reportes descargados desde la plataforma.

El objetivo principal es **analizar el uso de aplicaciones y reportes de usuarios**, integrando distintas fuentes de datos.

---

## 📁 Archivos del Proceso

El flujo del procesamiento se compone de los siguientes scripts:

- **`match_usuarios.py`**  
  Realiza el cruce (matching) de usuarios entre las distintas fuentes de datos disponibles.

- **`separar_usuarios.py`**  
  Filtra y separa los usuarios según los criterios definidos para el análisis.

- **`uso_plataforma.py`**  
  Genera métricas e informes relacionados con el uso de la plataforma por usuario.

---

## 📥 Insumos del Proceso

### 🔹 Insumos – Paso 1

Estos archivos deben descargarse previamente desde la plataforma antes de ejecutar el proceso.

#### 1️⃣ Archivo de descarga de usuarios
- **Archivo:** `User_Download_File.xlsx`  
- **Ruta de descarga:**  
  `Directorio > Usuarios > Descargar usuarios`

---

#### 2️⃣ Logs de uso de usuarios
- **Archivo:** `users_logs_File.xlsx`  
- **Ruta de descarga:**  
  `Informes > Denuncias de usuarios > Uso de las apps > Descargar`

---

#### 3️⃣ Reporte de Gemini por usuario
- **Archivo:** `Gemini_User_Reports_File.xlsx`  
- **Ruta de descarga:**  
  `IA Generativa > Informes de Gemini > Uso a nivel de los usuarios > Descargar tabla`

---
