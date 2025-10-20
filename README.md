# 📊 Automatización de Reporte de Productos con PIX Studio

## 🧩 Descripción del proyecto
Este proyecto automatiza el flujo completo de **extracción de datos desde SQL Server**, **creación y llenado de un reporte Excel (.xlsx)**, y finalmente **envía el archivo a través de un formulario web (Jotform)**.  

Está desarrollado íntegramente en **PIX Studio**, utilizando actividades de **Excel Interop**, **Selenium**, **SQL**, y control de flujo avanzado (**Try/Catch/Finally**, validaciones y logs).

---

## 🚀 Flujo general del proceso

### 1️⃣ Inicialización de variables
Se definen las rutas dinámicas para crear el archivo Excel con un nombre único basado en la fecha actual.

```csharp
excelPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Reportes", "Reporte_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx")
excelFileInfo = new FileInfo(excelPath)
```

👉 **Resultado:**  
Archivo generado automáticamente en la subcarpeta `Reportes` del proyecto, con un formato:
```
Reporte_2025-10-19.xlsx
```

---

### 2️⃣ Creación del archivo Excel
La creación y escritura del archivo se realiza con las **actividades nativas de Excel Interop** de PIX.  
Estas permiten crear un nuevo libro directamente al escribir datos por primera vez.  

El flujo verifica si el archivo existe:
```csharp
Si (!System.IO.Directory.Exists(excelFileInfo.DirectoryName))
```
➡️ Crea la carpeta `Reportes` si no existe.

Luego, las actividades **“Escribir en hoja”** generan automáticamente el archivo `Reporte_YYYY-MM-DD.xlsx` al escribir las tablas.

---

### 3️⃣ Extracción de información desde SQL Server
La base de datos contiene la tabla:

| Campo              | Tipo              |
|--------------------|-------------------|
| idProducto         | int               |
| Nombre             | nvarchar(255)     |
| Categoria          | nvarchar(100)     |
| Descripcion        | nvarchar(max)     |
| Precio             | decimal(10,2)     |
| FechaCreacion      | datetime          |
| FechaActualizacion | datetime          |

#### 🟢 Consulta de productos detallada
```sql
SELECT idProducto, Nombre, Categoria, Descripcion, Precio, FechaCreacion
FROM dbo.Productos;
```

#### 🟣 Consulta de resumen (total, promedios, por categoría)
```sql
SELECT 'Total de productos' AS Descripcion, COUNT(*) AS Valor1, NULL AS Valor2, 1 AS Orden
FROM dbo.Productos
UNION ALL
SELECT 'Precio promedio general', ROUND(AVG(Precio), 2), NULL, 2
FROM dbo.Productos
UNION ALL
SELECT '---', NULL, NULL, 3
UNION ALL
SELECT 'Categoría', 'Cantidad', 'Precio promedio', 4
UNION ALL
SELECT Categoria, COUNT(*) AS Valor1, ROUND(AVG(Precio), 2) AS Valor2, 5
FROM dbo.Productos
GROUP BY Categoria
ORDER BY Orden;
```

👉 Esta consulta devuelve una sola tabla consolidada con todos los indicadores requeridos para el resumen.

---

### 4️⃣ Llenado del Excel
Las actividades utilizadas:

- **“Escribir en hoja”** → Escribe los datos obtenidos de SQL directamente en el archivo Excel.  
- **Hoja1** → `vTablaProductos` (lista completa de productos).  
- **Hoja2** → `vTablaResumen` (indicadores agregados).  

Cada bloque comienza en la celda **A1** y crea automáticamente el archivo si no existe.

⚠️ Entre cada escritura se recomienda agregar una espera (`Esperar 2000 ms`) para evitar bloqueos temporales de acceso al archivo.

---

### 5️⃣ Formateo y extracción de fecha
Se extrae la fecha desde el nombre del archivo Excel para mostrarla en formato `MMddyyyy`:

```csharp
fechaFormateada = System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Match(Convert.ToString(excelPath), @"\d{4}-\d{2}-\d{2}").Value, @"(\d{4})-(\d{2})-(\d{2})", "$2$3$1")
```

Ejemplo:
```
Reporte_2025-10-19.xlsx → 10192025
```

---

### 6️⃣ Envío del reporte mediante formulario web
Se utilizó **Selenium en PIX** con las siguientes acciones:

#### 1. Iniciar navegador
Abre el formulario de Jotform:
```
https://form.jotform.com/252916633634057
```

#### 2. Llenar campos del formulario
- **Nombre:** `"Daniel"`
- **Fecha:** `fechaFormateada`
- **Upload:** ruta del archivo Excel (`excelPath`)

---

### 7️⃣ Scroll y clic controlado
Para garantizar que el botón de envío sea visible antes de hacer clic:
- Se usa la acción `ScrollIntoView` en el XPath del botón.
- Luego, un **If (vElementEnviar != "")** valida su existencia antes del clic.

En caso de que el botón no esté disponible:
```plaintext
No se completó el envío del formulario
```
Y se lanza una excepción personalizada.

---

### 8️⃣ Control de errores (Try/Catch/Finally)
Todas las secciones críticas (SQL, Excel, Selenium) están encapsuladas con manejo de errores robusto.

Ejemplo de captura:
```csharp
"Error al llenar el formulario. Detalle: " + Convert.ToString(exc.Message)
```

También se registran logs de ejecución con:
```plaintext
Trace Info → "Se envió correctamente el formulario."
Trace Error → "Error al llenar el formulario: " + exc.Message
```

---

## 🧠 Flujo final resumido

1. Inicializa variables y ruta del archivo.  
2. Crea carpeta si no existe.  
3. Consulta SQL de productos y resumen.  
4. Llena las dos hojas del Excel (Hoja1 y Hoja2).  
5. Genera la fecha formateada desde el nombre del archivo.  
6. Abre navegador con Selenium.  
7. Llena el formulario y sube el archivo Excel.  
8. Hace scroll y clic controlado en Enviar.  
9. Escribe logs y maneja excepciones de ejecución.  

---

## ⚙️ Requerimientos

| Dependencia | Descripción |
|--------------|-------------|
| Microsoft Excel | Requerido para las operaciones Interop. |
| PIX Studio | Entorno principal de automatización. |
| Selenium | Control de navegador. |
| SQL Server | Fuente de datos principal. |

---

## ✅ Resultado final
El bot genera, llena y sube automáticamente un archivo Excel con la siguiente estructura:

### Hoja1 – Productos
| idProducto | Nombre | Categoria | Descripcion | Precio | FechaCreacion |
|-------------|---------|------------|--------------|---------|----------------|
| 1 | Producto A | electronics | ... | 199.90 | 2025-10-17 |
| ... | ... | ... | ... | ... | ... |

### Hoja2 – Resumen
| Descripcion | Valor1 | Valor2 |
|--------------|---------|--------|
| Total de productos | 20 | |
| Precio promedio general | 162.05 | |
| ... | ... | ... |

---

## 🧾 Créditos
**Desarrollado por:** Daniel Ortiz Correa  
**Herramientas:** PIX Studio, SQL Server, Excel Interop, Selenium  

---

## 🧠 Notas finales
- El Excel se genera directamente al escribir los datos (sin scripts externos).  
- Cada ejecución crea un archivo nuevo con la fecha actual.  
- El flujo es escalable y adaptable a otros formularios o estructuras SQL.  

---

> 💡 *Este flujo está diseñado para ser estable, reutilizable y completamente automatizado. Ideal para procesos de reporting o carga de evidencias diarias.*
