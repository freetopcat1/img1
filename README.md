# 📋 Outlook Add-in: Email a Microsoft Planner

Complemento de Outlook que permite crear tareas en Microsoft Planner directamente desde un correo electrónico seleccionado, incluyendo los archivos adjuntos.

---

## 🗂 Estructura del Proyecto

```
outlook-planner-addin/
├── manifest.xml              ← Manifiesto del add-in (se registra en Outlook)
├── package.json              ← Dependencias y scripts de desarrollo
├── README.md                 ← Este archivo
└── src/
    ├── taskpane.html         ← Interfaz principal del complemento
    ├── taskpane.js           ← Toda la lógica (Office.js + Graph API)
    ├── commands.html         ← Archivo de comandos requerido por el manifiesto
    └── auth.html             ← Página de redirección para autenticación MSAL
```

---

## ✅ Prerrequisitos

| Herramienta | Versión mínima |
|-------------|---------------|
| Node.js     | 16 o superior |
| Outlook     | Microsoft 365 (Web, Windows o Mac) |
| Cuenta M365 | Con licencia de Planner |

---

## 🔐 Paso 1 — Registrar la aplicación en Azure Active Directory

1. Ve a [portal.azure.com](https://portal.azure.com) e inicia sesión con tu cuenta de administrador.
2. Navega a **Azure Active Directory → Registros de aplicaciones → Nueva registración**.
3. Completa el formulario:
   - **Nombre:** `Outlook Planner Add-in` (o el que prefieras)
   - **Tipos de cuenta:** *Cuentas en este directorio organizacional únicamente*
   - **URI de redirección:** Selecciona **Aplicación de página única (SPA)** e ingresa:
     ```
     https://TU_DOMINIO/src/auth.html
     ```
     > Para pruebas locales usa: `https://localhost:3000/src/auth.html`
4. Haz clic en **Registrar**.

### Anota estos valores (los necesitarás en el paso 3):
- **Application (client) ID** → `TU_CLIENT_ID_AQUI`
- **Directory (tenant) ID** → `TU_TENANT_ID_AQUI`

### Agregar permisos de API

En tu app registrada → **Permisos de API → Agregar un permiso → Microsoft Graph → Permisos delegados**:

| Permiso | Para qué se usa |
|---------|-----------------|
| `Group.Read.All` | Listar los grupos de Office 365 del usuario |
| `GroupMember.Read.All` | Ver los miembros del grupo seleccionado |
| `Tasks.ReadWrite` | Crear y actualizar tareas en Planner |
| `Files.ReadWrite.All` | Subir adjuntos al SharePoint del grupo |
| `Mail.Read` | Leer el correo seleccionado (respaldo) |

> ⚠️ **Importante:** Haz clic en **Conceder consentimiento de administrador** después de agregar los permisos.

---

## ⚙️ Paso 2 — Configurar el proyecto

### Clona o copia los archivos
```bash
cd outlook-planner-addin
npm install
```

### Actualiza las URLs en `manifest.xml`
Reemplaza **todas** las ocurrencias de `YOUR_DOMAIN` con tu dominio real o `localhost:3000`:
```xml
<!-- Ejemplo para producción -->
<SourceLocation DefaultValue="https://miempresa.com/src/taskpane.html"/>

<!-- Ejemplo para desarrollo local -->
<SourceLocation DefaultValue="https://localhost:3000/src/taskpane.html"/>
```

Las líneas a modificar son:
- `<IconUrl>`
- `<HighResolutionIconUrl>`
- `<SupportUrl>`
- `<bt:Url id="Commands.Url">`
- `<bt:Url id="Taskpane.Url">`
- Todas las `<bt:Image ... DefaultValue="...">`

---

## 🔑 Paso 3 — Configurar las credenciales en el código

Abre `src/taskpane.js` y `src/auth.html`, y reemplaza:

```javascript
// En src/taskpane.js — sección CONFIG (línea ~20)
const CONFIG = {
  clientId : 'TU_CLIENT_ID_AQUI',    // ← pega tu Application ID
  tenantId : 'TU_TENANT_ID_AQUI',    // ← pega tu Directory (tenant) ID
  ...
};
```

```javascript
// En src/auth.html — script al final de la página
clientId: 'TU_CLIENT_ID_AQUI',       // ← mismo valor que arriba
```

---

## 🚀 Paso 4 — Ejecutar en desarrollo local

```bash
# Instalar certificado SSL local (necesario para Office Add-ins)
npm start
```

El servidor arrancará en `https://localhost:3000`.

> Si aparece un error de certificado, ejecuta primero:
> ```bash
> npx office-addin-dev-certs install
> ```

---

## 📦 Paso 5 — Cargar el complemento en Outlook

### Opción A — Carga de prueba (sideload) en Outlook Web
1. Ve a [outlook.office.com](https://outlook.office.com)
2. Abre cualquier correo → haz clic en los **tres puntos (···)** → **Obtener complementos**
3. En la ventana de complementos → **Mis complementos → Agregar un complemento personalizado → Agregar desde archivo**
4. Selecciona el archivo `manifest.xml`

### Opción B — Sideload en Outlook Desktop (Windows)
1. Outlook → **Archivo → Administrar complementos** → (se abre en el navegador)
2. Repite los pasos de la Opción A

### Opción C — Despliegue centralizado (producción)
1. Centro de administración de Microsoft 365 → **Configuración → Aplicaciones integradas**
2. Sube el `manifest.xml` y asigna usuarios

---

## 🖥️ Cómo usar el complemento

1. Abre Outlook y selecciona un correo electrónico.
2. En la cinta de opciones verás el botón **📋 Email a Planner** — haz clic.
3. El panel lateral se abrirá. Haz clic en **Iniciar Sesión con Microsoft**.
4. Completa el formulario:
   - Selecciona el **grupo de Office 365** del que eres propietario
   - Selecciona el **Plan de Planner** dentro de ese grupo
   - Revisa el asunto, cuerpo y adjuntos del correo
5. Haz clic en **✅ Crear Tarea Planner**.
6. El complemento:
   - Crea la tarea con el asunto como título
   - Agrega el cuerpo del correo como descripción
   - Sube los archivos adjuntos al SharePoint del grupo
   - Vincula los archivos como referencias en la tarea de Planner
7. Haz clic en **✖ Cancelar** para cerrar sin crear nada.

---

## 🏗️ Despliegue en producción

### Con Azure Static Web Apps (recomendado)
```bash
# 1. Instala Azure CLI
# 2. Crea una Static Web App
az staticwebapp create --name outlook-planner-addin --resource-group MiGrupo

# 3. Despliega
az staticwebapp deploy --app-location "." --output-location "."
```

### Con cualquier servidor web (Nginx, IIS, Apache)
Simplemente copia todos los archivos a una carpeta pública del servidor con HTTPS habilitado.

> ⚠️ **HTTPS es obligatorio** — Office Add-ins sólo funcionan sobre HTTPS.

---

## ⚠️ Notas importantes

### Compatibilidad de adjuntos
La API `getAttachmentContentAsync()` requiere **Mailbox requirement set 1.8** (Outlook 2019 / Microsoft 365). En versiones anteriores, la tarea se creará correctamente pero sin adjuntos subidos a SharePoint.

### Límite de descripción de Planner
Planner limita la descripción de una tarea a **1 024 caracteres**. Correos más largos serán truncados en la descripción, pero el cuerpo completo del correo original sigue accesible en Outlook.

### Permisos de SharePoint
Para que los archivos se suban correctamente, el usuario debe tener acceso de escritura al sitio de SharePoint asociado al grupo de Office 365 seleccionado.

---

## 🐛 Solución de problemas comunes

| Problema | Solución |
|----------|----------|
| "Error de autenticación" | Verifica que `clientId` y `tenantId` estén correctos |
| "No eres propietario de ningún grupo" | Asegúrate de ser **propietario** (no sólo miembro) del grupo |
| "Este grupo no tiene planes" | Crea al menos un plan en Planner para ese grupo |
| El popup de login se bloquea | Habilita popups para tu dominio en el navegador |
| Adjuntos no se suben | Verifica el permiso `Files.ReadWrite.All` y que la versión de Outlook sea ≥ 2019 |

---

## 📄 Licencia

MIT — Uso libre para proyectos personales y empresariales.
