/* =============================================================
   taskpane.js  —  Outlook Add-in: Email a Microsoft Planner
   =============================================================

   REQUISITOS PREVIOS:
   1. Registra una app en Azure Portal (portal.azure.com)
   2. En "Autenticación" habilita "Single-page application" con la
      redirectUri de tu dominio + /src/auth.html
   3. En "Permisos de API" agrega los permisos delegados de Graph
      indicados abajo y otorga consentimiento de administrador.
   4. Copia el "Application (client) ID" y "Directory (tenant) ID"
      en la sección CONFIG de este archivo.
============================================================= */

'use strict';

/* =============================================================
   ① CONFIGURACIÓN  —  ¡ACTUALIZA ESTOS VALORES!
============================================================= */
const CONFIG = {
  clientId : '3214aa21-198e-472c-b430-3dbf8d402dd4',       // App Registration > Application (client) ID
  tenantId : 'd239009e-8fbd-445b-9cb0-2b37507f6006',       // App Registration > Directory (tenant) ID
  // redirectUri apunta a la página auth.html que maneja el popup de MSAL
  get redirectUri() {
    const base = window.location.href.substring(0, window.location.href.lastIndexOf('/'));
    return `${base}/auth.html`;
  },
  // Permisos de Microsoft Graph necesarios:
  //   Group.Read.All           → listar grupos
  //   GroupMember.Read.All     → listar miembros del grupo
  //   Tasks.ReadWrite          → crear tareas en Planner
  //   Files.ReadWrite.All      → subir adjuntos a SharePoint del grupo
  //   Mail.Read                → (respaldo) leer correo por REST
  
	scopes: [
	  "User.Read",
	  "Group.Read.All",
	  "GroupMember.Read.All",
	  "Planner.ReadWrite",
	  "Files.ReadWrite.All",
	  "Mail.Read"
	]
};

/* =============================================================
   ② ESTADO DE LA APLICACIÓN
============================================================= */
let msalInstance   = null;
let accessToken    = null;
let emailItem      = null;   // Office.context.mailbox.item
let emailSubject   = '';
let emailBody      = '';
let emailBodyHtml  = '';
let emailAttachments = [];   // array de objetos AttachmentDetails de Office.js

/* =============================================================
   ③ PUNTO DE ENTRADA: Office.onReady
============================================================= */
// ========== PUNTO DE ENTRADA ==========
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) return;

  try {
    // 1. Obtener token silenciosamente usando la sesión de Office
    await acquireTokenWithOfficeAuth();

    // 2. Cargar datos del correo
    await loadEmailData();

    // 3. Configurar eventos UI y mostrar contenido principal
    setupEventListeners();
    showMainContent();

    // 4. Cargar grupos del usuario
    await loadOwnedGroups();
	
	//5. hacer Auth Section = False
	document.getElementById('authSection').style.display = 'none';
	

  } catch (err) {
    showGlobalError('No se pudo autenticar automáticamente: ' + err.message);
  }
});

alert("OK CARGADO");

// ========== OBTENER TOKEN CON Office.auth ==========
async function acquireTokenWithOfficeAuth() {
  return new Promise((resolve, reject) => {
    Office.auth.getAccessToken({
      allowSignInPrompt: false,    // No mostrar diálogo si no hay sesión
      allowConsentPrompt: false,   // No pedir consentimiento (debe estar pre-aprobado)
      forMSGraphAccess: true,      // Indica que es para Microsoft Graph
      authChallenge: (challenge) => {
        // Si ocurre un desafío de autenticación (ej. falta consentimiento)
        // puedes manejarlo aquí, pero por ahora rechazamos
        reject(new Error('Se requiere consentimiento o autenticación adicional.'));
      }
    }, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        accessToken = asyncResult.value;
        resolve();
      } else {
        reject(asyncResult.error);
      }
    });
  });
}

// El resto de funciones (loadEmailData, loadOwnedGroups, graphFetch, etc.)
// se mantienen igual, porque ya usan accessToken.
// Solo asegúrate de que `graphFetch` verifique accessToken como antes.

/* =============================================================
   ⑤ LECTURA DE DATOS DEL CORREO (Office.js)
============================================================= */
async function loadEmailData() {
  emailItem = Office.context.mailbox.item;

  // Asunto
  emailSubject = emailItem.subject || '(Sin asunto)';
  setVal('subjectValue', emailSubject);

  // Cuerpo (texto plano)
  emailBody = await new Promise((resolve) => {
    emailItem.body.getAsync('text', { asyncContext: null }, (res) => {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : '');
    });
  });

  // Cuerpo HTML (para adjuntarlo como descripción extendida si se requiere)
  emailBodyHtml = await new Promise((resolve) => {
    emailItem.body.getAsync('html', { asyncContext: null }, (res) => {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : '');
    });
  });

  // Mostrar vista previa del cuerpo (máx. 600 chars)
  const preview = emailBody.trim();
  setVal('bodyValue', preview ? preview.substring(0, 600) + (preview.length > 600 ? '…' : '') : '(Sin cuerpo)');

  // Adjuntos (sólo los que no son inline)
  emailAttachments = (emailItem.attachments || []).filter(a => !a.isInline);
  renderAttachments(emailAttachments);
}

/* =============================================================
   ⑥ GRAPH API — HELPERS
============================================================= */
const GRAPH = 'https://graph.microsoft.com/v1.0';

async function graphFetch(method, endpoint, body = null, extraHeaders = {}) {
  if (!accessToken) throw new Error('No hay token de acceso. Autentícate primero.');

  const headers = {
    'Authorization': `Bearer ${accessToken}`,
    'Content-Type' : 'application/json',
    ...extraHeaders,
  };

  const opts = { method, headers };
  if (body !== null && typeof body === 'object' && !(body instanceof ArrayBuffer) && !ArrayBuffer.isView(body)) {
    opts.body = JSON.stringify(body);
  } else if (body !== null) {
    opts.body = body;
    headers['Content-Type'] = extraHeaders['Content-Type'] || 'application/octet-stream';
  }

  const resp = await fetch(GRAPH + endpoint, opts);

  if (resp.status === 204) return null;           // No Content
  if (!resp.ok) {
    let msg = `Error ${resp.status}`;
    try { const e = await resp.json(); msg = e.error?.message || msg; } catch (_) {}
    throw new Error(msg);
  }
  return resp.json();
}

const graphGet   = (ep, hdrs)      => graphFetch('GET',   ep, null, hdrs);
const graphPost  = (ep, body)      => graphFetch('POST',  ep, body);
const graphPatch = (ep, body, etag) => graphFetch('PATCH', ep, body, etag ? { 'If-Match': etag } : {});
const graphPut   = (ep, data, ct)  => graphFetch('PUT',   ep, data, { 'Content-Type': ct || 'application/octet-stream' });

/* =============================================================
   ⑦ GRUPOS DE OFFICE 365 (propios)
============================================================= */
async function loadOwnedGroups() {
  setLoading('Cargando tus grupos de Office 365…');
  try {
    // /me/ownedObjects/microsoft.graph.group → sólo objetos "grupo" que posee el usuario
    const data = await graphGet(
      '/me/ownedObjects/microsoft.graph.group?$select=id,displayName,mail,groupTypes&$top=100'
    );

    // Filtrar sólo grupos Unified (Office 365 groups / Microsoft 365 groups)
    const unified = (data.value || []).filter(
      g => Array.isArray(g.groupTypes) && g.groupTypes.includes('Unified')
    );

    renderGroups(unified);
    clearLoading();
  } catch (err) {
    showError('Error cargando grupos: ' + err.message);
  }
}

/* =============================================================
   ⑧ DATOS DEL GRUPO (miembros + planes de Planner)
============================================================= */
async function loadGroupData(groupId) {
  setLoading('Cargando datos del grupo…');
  try {
    const [membersData, plansData] = await Promise.all([
      graphGet(`/groups/${groupId}/members?$select=id,displayName,mail,userPrincipalName&$top=100`),
      graphGet(`/groups/${groupId}/planner/plans?$select=id,title`),
    ]);

    renderMembers(membersData.value || []);
    renderPlans(plansData.value || []);
    clearLoading();
  } catch (err) {
    showError('Error cargando datos del grupo: ' + err.message);
  }
}

/* =============================================================
   ⑨ CREACIÓN DE TAREA EN PLANNER + ADJUNTOS
============================================================= */
async function createPlannerTask() {
  const groupId = document.getElementById('groupSelect').value;
  const planId  = document.getElementById('planSelect').value;

  if (!groupId) { showError('Selecciona un grupo de Office 365.'); return; }
  if (!planId)  { showError('Selecciona un plan de Planner.'); return; }

  disableButtons(true);

  try {
    /* ── 1. Obtener primer bucket del plan ── */
    setLoading('Obteniendo buckets del plan…');
    let bucketId = null;
    try {
      const buckets = await graphGet(`/planner/plans/${planId}/buckets?$select=id,name`);
      if (buckets.value && buckets.value.length > 0) {
        bucketId = buckets.value[0].id;
      }
    } catch (_) { /* continúa sin bucket */ }

    /* ── 2. Crear la tarea ── */
    setLoading('Creando tarea en Planner…');
    const taskPayload = {
      planId : planId,
      title  : emailSubject || 'Tarea desde correo',
    };
    if (bucketId) taskPayload.bucketId = bucketId;

    const task   = await graphPost('/planner/tasks', taskPayload);
    const taskId = task.id;

    /* ── 3. Obtener detalles de la tarea para conseguir el eTag ── */
    setLoading('Actualizando detalles de la tarea…');
    const details  = await graphGet(`/planner/tasks/${taskId}/details`);
    const etag     = details['@odata.etag'];

    /* ── 4. Subir archivos adjuntos al SharePoint del grupo ── */
    const references = {};

    if (emailAttachments.length > 0) {
      for (let i = 0; i < emailAttachments.length; i++) {
        const att = emailAttachments[i];
        setLoadingProgress(
          `Adjuntando archivo ${i + 1} de ${emailAttachments.length}: ${att.name}…`,
          Math.round(((i + 1) / emailAttachments.length) * 100)
        );

        try {
          const content = await getAttachmentContent(att.id);
          if (content) {
            /* Subir al drive del grupo en SharePoint:
               /groups/{id}/drive/root:/PlannerTasks/{taskId}/{fileName}:/content  */
            const uploadPath = `/groups/${groupId}/drive/root:/PlannerTasks/${taskId}/${sanitizeFileName(att.name)}:/content`;
            const uploaded   = await graphPut(uploadPath, content.data, att.contentType || 'application/octet-stream');

            /* Agregar como referencia externa en Planner */
            const fileUrl = uploaded.webUrl;
            /* La clave debe ser la URL codificada (Planner usa URL como clave del objeto) */
            const encodedUrl = encodeURIComponent(fileUrl);
            references[encodedUrl] = {
              '@odata.type'   : 'microsoft.graph.plannerExternalReference',
              'alias'         : att.name,
              'type'          : 'Other',
              'previewPriority': ' !',
            };
          } else {
            console.warn(`No se pudo obtener contenido del adjunto: ${att.name}`);
          }
        } catch (attachErr) {
          console.error(`Error procesando adjunto "${att.name}":`, attachErr);
          // Continúa con el siguiente adjunto sin abortar
        }
      }
      clearProgress();
    }

    /* ── 5. Actualizar detalles: descripción + referencias ── */
    setLoading('Guardando descripción y referencias…');
    const detailsPayload = {
      // Planner admite hasta 1 024 chars en description
      description: truncate(emailBody, 1024),
    };
    if (Object.keys(references).length > 0) {
      detailsPayload.references = references;
    }

    await graphPatch(`/planner/tasks/${taskId}/details`, detailsPayload, etag);

    /* ── 6. Éxito ── */
    showSuccess(`✅ Tarea "${emailSubject}" creada en Planner. El panel se cerrará en 3 segundos…`);
    setTimeout(() => {
      try { Office.addin.hide(); } catch (_) { /* navegadores que no lo soporten */ }
    }, 3000);

  } catch (err) {
    showError('Error al crear la tarea: ' + err.message);
    disableButtons(false);
  }
}

/* =============================================================
   ⑩ OBTENER CONTENIDO DE UN ADJUNTO (Office.js ≥ 1.8)
============================================================= */
function getAttachmentContent(attachmentId) {
  return new Promise((resolve) => {
    /* getAttachmentContentAsync requiere Mailbox 1.8 */
    if (typeof emailItem.getAttachmentContentAsync !== 'function') {
      resolve(null);
      return;
    }

    emailItem.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error('getAttachmentContentAsync falló:', result.error);
        resolve(null);
        return;
      }

      const val = result.value;
      /* format === Office.MailboxEnums.AttachmentContentFormat.Base64 */
      try {
        const binary = atob(val.content);
        const bytes  = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        resolve({ data: bytes.buffer, format: val.format });
      } catch (e) {
        console.error('Error decodificando adjunto:', e);
        resolve(null);
      }
    });
  });
}

/* =============================================================
   ⑪ RENDERIZADO DE LA UI
============================================================= */
function renderGroups(groups) {
  const sel = document.getElementById('groupSelect');
  sel.innerHTML = '<option value="">-- Selecciona un grupo --</option>';

  if (groups.length === 0) {
    sel.innerHTML = '<option value="">No eres propietario de ningún grupo</option>';
    return;
  }

  groups.forEach(g => {
    const opt     = document.createElement('option');
    opt.value     = g.id;
    opt.textContent = g.displayName;
    sel.appendChild(opt);
  });
}

function renderPlans(plans) {
  const container = document.getElementById('planContainer');
  const sel       = document.getElementById('planSelect');
  const btn       = document.getElementById('createBtn');

  container.style.display = 'block';
  sel.innerHTML = '<option value="">-- Selecciona un plan --</option>';

  if (plans.length === 0) {
    sel.innerHTML = '<option value="" disabled>Este grupo no tiene planes en Planner</option>';
    btn.disabled  = true;
    return;
  }

  plans.forEach(p => {
    const opt       = document.createElement('option');
    opt.value       = p.id;
    opt.textContent = p.title;
    sel.appendChild(opt);
  });

  // Habilitar el botón sólo cuando se elija un plan
  sel.addEventListener('change', () => {
    btn.disabled = !sel.value;
  });
}

function renderMembers(members) {
  const list  = document.getElementById('membersList');
  const card  = document.getElementById('membersCard');
  const badge = document.getElementById('memberBadge');

  list.innerHTML = '';
  badge.textContent = members.length;
  card.style.display = 'block';

  if (members.length === 0) {
    list.innerHTML = '<li class="empty-note">No se encontraron miembros.</li>';
    return;
  }

  members.forEach(m => {
    const li  = document.createElement('li');
    li.className = 'member-item';
    li.innerHTML = `
      <span class="member-name">${esc(m.displayName || '')}</span>
      <span class="member-email">${esc(m.mail || m.userPrincipalName || '')}</span>
    `;
    list.appendChild(li);
  });
}

function renderAttachments(attachments) {
  const list  = document.getElementById('attachmentsList');
  const badge = document.getElementById('attachBadge');

  badge.textContent = attachments.length;
  list.innerHTML    = '';

  if (attachments.length === 0) {
    list.innerHTML = '<li class="empty-note">Sin archivos adjuntos</li>';
    return;
  }

  attachments.forEach(a => {
    const li      = document.createElement('li');
    li.className  = 'attachment-item';
    li.innerHTML  = `
      <span class="attachment-icon">${fileIcon(a.name)}</span>
      <span class="attachment-name" title="${esc(a.name)}">${esc(a.name)}</span>
      <span class="attachment-size">${formatSize(a.size)}</span>
    `;
    list.appendChild(li);
  });
}

/* =============================================================
   ⑫ EVENTOS DE LA UI
============================================================= */
function setupEventListeners() {
  /* Botón de autenticación */
  document.getElementById('authenticateBtn').addEventListener('click', async () => {
    const authErr = document.getElementById('authError');
    authErr.style.display = 'none';
    try {
      document.getElementById('authenticateBtn').textContent = 'Autenticando…';
      document.getElementById('authenticateBtn').disabled = true;

      accessToken = await authenticate();

      // Cambiar a contenido principal
      document.getElementById('authSection').style.display = 'none';
      showMainContent();
      await loadOwnedGroups();

    } catch (err) {
      authErr.textContent = 'Error de autenticación: ' + err.message;
      authErr.style.display = 'flex';
      document.getElementById('authenticateBtn').textContent = 'Reintentar';
      document.getElementById('authenticateBtn').disabled = false;
    }
  });

  /* Cambio de grupo */
  document.getElementById('groupSelect').addEventListener('change', async (e) => {
    const gid = e.target.value;

    // Resetear secciones dependientes
    document.getElementById('membersCard').style.display  = 'none';
    document.getElementById('planContainer').style.display = 'none';
    document.getElementById('createBtn').disabled         = true;

    if (!gid) return;
    await loadGroupData(gid);
  });

  /* Botón Crear Tarea */
  document.getElementById('createBtn').addEventListener('click', createPlannerTask);

  /* Botón Cancelar */
  document.getElementById('cancelBtn').addEventListener('click', () => {
    try {
      // Office.addin.hide() oculta el panel de tareas (Microsoft 365 / Office 2019+)
      Office.addin.hide();
    } catch (_) {
      // Fallback para versiones antiguas: limpiar el contenido
      document.getElementById('mainContent').innerHTML =
        '<div style="padding:24px;text-align:center;color:#605e5c">'+
        '<p>Panel cerrado.</p><p style="font-size:12px;margin-top:8px">'+
        'Puedes cerrar este panel con el botón ✕ de la barra lateral.</p></div>';
    }
  });
}

/* =============================================================
   ⑬ HELPERS DE ESTADO / UI
============================================================= */
function showAuthSection() {
  document.getElementById('authSection').style.display    = 'flex';
  document.getElementById('mainContent').style.display   = 'none';
}

function showMainContent() {
  document.getElementById('authSection').style.display  = 'none';
  const mc = document.getElementById('mainContent');
  mc.style.display = 'flex';
  mc.style.flexDirection = 'column';
}

function setLoading(text) {
  const bar  = document.getElementById('loadingMessage');
  const span = document.getElementById('loadingText');
  span.textContent  = text;
  bar.style.display = 'flex';
  document.getElementById('errorMessage').style.display   = 'none';
  document.getElementById('successMessage').style.display = 'none';
}

function setLoadingProgress(text, pct) {
  setLoading(text);
  const wrap  = document.getElementById('progressWrap');
  const fill  = document.getElementById('progressFill');
  const label = document.getElementById('progressLabel');
  wrap.style.display  = 'block';
  fill.style.width    = pct + '%';
  label.textContent   = pct + '%';
}

function clearLoading() {
  document.getElementById('loadingMessage').style.display = 'none';
}

function clearProgress() {
  document.getElementById('progressWrap').style.display = 'none';
}

function showError(msg) {
  clearLoading();
  clearProgress();
  const bar = document.getElementById('errorMessage');
  bar.textContent   = msg;
  bar.style.display = 'flex';
}

function showSuccess(msg) {
  clearLoading();
  clearProgress();
  const bar = document.getElementById('successMessage');
  bar.textContent   = msg;
  bar.style.display = 'flex';
  document.getElementById('errorMessage').style.display = 'none';
}

function showGlobalError(msg) {
  document.body.innerHTML =
    `<div style="padding:24px;color:#d13438;font-family:Segoe UI,sans-serif">
      <strong>Error</strong><br/><span>${esc(msg)}</span>
     </div>`;
}

function disableButtons(disabled) {
  document.getElementById('createBtn').disabled = disabled;
  document.getElementById('cancelBtn').disabled = disabled;
}

function setVal(id, text) {
  const el = document.getElementById(id);
  if (el) el.textContent = text;
}

/* =============================================================
   ⑭ UTILIDADES
============================================================= */
function esc(str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function truncate(str, max) {
  if (!str) return '';
  return str.length <= max ? str : str.substring(0, max - 1) + '…';
}

function formatSize(bytes) {
  if (!bytes || bytes === 0) return '';
  if (bytes < 1024)          return bytes + ' B';
  if (bytes < 1024 * 1024)   return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function sanitizeFileName(name) {
  // Elimina caracteres no permitidos en rutas de SharePoint
  return name.replace(/[#%*:<>?/\\|"]/g, '_');
}

function fileIcon(name) {
  const ext = (name || '').split('.').pop().toLowerCase();
  const icons = {
    pdf: '📄', doc: '📝', docx: '📝', xls: '📊', xlsx: '📊',
    ppt: '📊', pptx: '📊', jpg: '🖼️', jpeg: '🖼️', png: '🖼️',
    gif: '🖼️', zip: '🗜️', rar: '🗜️', txt: '📃', csv: '📊',
    mp4: '🎬', mp3: '🎵', msg: '📧',
  };
  return icons[ext] || '📎';
}
