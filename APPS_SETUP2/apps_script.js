// ══════════════════════════════════════════════════════════════
// APPS SCRIPT — BACKEND AGRICOLA GUAPA SAS
// ══════════════════════════════════════════════════════════════

const SHEET_ID = '13OQgh4EZl0ArkAcki16Jh_8Ia4VpcV10YiUiuOlQyy0';

function getSheet(nombre) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    if (nombre === 'Anomalias') {
      sheet.appendRow(['ID','Fecha','Hora','Operador','Equipo','Disponibilidad',
        'Sede','Lote','Area','Sistema','Subsistema','Parte','Modo Falla','Prioridad',
        'Falla','Observaciones','Estado','Mecanismo','Tecnico',
        'Inicio Reparacion','Fin Reparacion','Obs Taller','Fecha Estado','Sincronizado']);
      sheet.getRange(1,1,1,24).setFontWeight('bold').setBackground('#1E40AF').setFontColor('#FFFFFF');
    } else if (nombre === 'Labores') {
      sheet.appendRow(['ID Turno','Fecha','Operador','Equipo','Area','Finca','Horometro Ini',
        'Fecha Cierre','Horometro Fin','ACPM','Hora Labor','Cod Labor','Desc Labor',
        'Clasificacion','Distribucion','Lote','Implemento','Cantidad','Unidad','Tipo Cosecha','Obs','Sincronizado']);
      sheet.getRange(1,1,1,21).setFontWeight('bold').setBackground('#065F46').setFontColor('#FFFFFF');
    }
  }
  return sheet;
}

function doPost(e) {
  try {
    // Aceptar JSON body o URLSearchParams (para compatibilidad con file://)
    let data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch(_) {
      const p = e.parameter;
      if (p && p.data) {
        const parsed = JSON.parse(p.data);
        data = {
          action:   p.action,
          registros: Array.isArray(parsed) ? parsed : [parsed],
          turnos:    Array.isArray(parsed) ? parsed : [parsed],
          registro:  Array.isArray(parsed) ? parsed[0] : parsed
        };
      } else {
        data = p || {};
      }
    }

    const action = data.action;

    // ── Guardar anomalía individual
    if (action === 'saveAnomalia') {
      const sheet = getSheet('Anomalias');
      const a = data.registro;
      const existe = getIdsExistentes(sheet);
      if (!existe.has(String(a.id))) {
        sheet.appendRow([a.id,a.fecha,a.hora,a.operador,a.equipo,a.disponibilidad,
          a.sede||'',a.lote||'',a.area||'',a.sistema||'',a.subsistema||'',a.parte||'',
          a.modo_falla||'',a.prioridad||'',a.falla||'',a.observaciones||'',
          a.estado||'Abierta',a.mecanismo||'',a.tecnico||'',
          a.fecha_ini_rep||'',a.fecha_fin_rep||'',a.obs_taller||'',
          a.fecha_estado||'',new Date().toISOString()]);
      }
      return jsonOk();
    }

    // ── Guardar lote de anomalías
    if (action === 'saveLote') {
      const sheet = getSheet('Anomalias');
      const registros = data.registros || [];
      const existe = getIdsExistentes(sheet);
      let nuevos = 0;
      registros.forEach(a => {
        if (!existe.has(String(a.id))) {
          sheet.appendRow([a.id,a.fecha,a.hora,a.operador,a.equipo,a.disponibilidad,
            a.sede||'',a.lote||'',a.area||'',a.sistema||'',a.subsistema||'',a.parte||'',
            a.modo_falla||'',a.prioridad||'',a.falla||'',a.observaciones||'',
            a.estado||'Abierta',a.mecanismo||'',a.tecnico||'',
            a.fecha_ini_rep||'',a.fecha_fin_rep||'',a.obs_taller||'',
            a.fecha_estado||'',new Date().toISOString()]);
          nuevos++;
        }
      });
      return jsonOk({nuevos: nuevos});
    }

    // ── Actualizar anomalía (taller)
    if (action === 'updateAnomalia') {
      const sheet = getSheet('Anomalias');
      const a = data.registro;
      const datos = sheet.getDataRange().getValues();
      for (let i = 1; i < datos.length; i++) {
        if (String(datos[i][0]) === String(a.id)) {
          sheet.getRange(i+1,17).setValue(a.estado||'');
          sheet.getRange(i+1,18).setValue(a.mecanismo||'');
          sheet.getRange(i+1,19).setValue(a.tecnico||'');
          sheet.getRange(i+1,20).setValue(a.fecha_ini_rep||'');
          sheet.getRange(i+1,21).setValue(a.fecha_fin_rep||'');
          sheet.getRange(i+1,22).setValue(a.obs_taller||'');
          sheet.getRange(i+1,23).setValue(a.fecha_estado||'');
          sheet.getRange(i+1,13).setValue(a.modo_falla||'');
          return jsonOk();
        }
      }
      return jsonOk({msg:'No encontrado'});
    }

    // ── Guardar turno
    if (action === 'saveTurnos' || action === 'saveTurno') {
      const sheet = getSheet('Labores');
      const turnos = data.turnos || (data.turno ? [data.turno] : []);
      const existe = getIdsExistentes(sheet);
      let nuevos = 0;
      turnos.forEach(t => {
        // ID único por turno: fecha + operador_codigo + equipo
        const tid = String(t.id || '') || ((t.fecha||'') + '_' + (t.operador||'') + '_' + (t.equipo||''));
        if (!existe.has(String(tid))) {
          const acts = t.actividades || [];
          if (acts.length === 0) {
            // Turno sin actividades — guardar fila vacía
            sheet.appendRow([tid, t.fecha, t.operador_nombre||t.operador, t.equipo_desc||t.equipo,
              t.area||'', t.finca||'', t.horometro_ini||'', t.fecha_cierre||'',
              t.horometro_fin||'', t.acpm||'',
              '','','','','','','','','', new Date().toISOString()]);
          } else {
            acts.forEach(a => {
              // Combinar implemento código + descripción en un solo campo
              const implText = a.implemento
                ? (a.implemento + (a.impl_desc ? ' - ' + a.impl_desc : ''))
                : '';
              sheet.appendRow([
                tid,
                t.fecha,
                t.operador_nombre || t.operador,   // Operador
                t.equipo_desc     || t.equipo,      // Equipo
                t.area            || '',             // Area
                a.finca || t.finca || '',            // Finca (actividad o turno)
                t.horometro_ini   || '',             // Horometro Ini
                t.fecha_cierre    || '',             // Fecha Cierre
                t.horometro_fin   || '',             // Horometro Fin
                t.acpm            || '',             // ACPM
                a.hora            || '',             // Hora Labor
                a.labor           || '',             // Cod Labor
                a.labor_desc      || '',             // Desc Labor
                a.clasificacion   || '',             // Clasificacion (de la labor)
                a.distribucion    || '',             // Distribucion (de la labor)
                a.lote            || '',             // Lote
                implText,                            // Implemento (código - descripción)
                a.cantidad        || '',             // Cantidad
                a.unidad          || '',             // Unidad
                a.tipo_cosecha    || '',             // Tipo Cosecha
                a.obs             || '',             // Obs
                new Date().toISOString()             // Sincronizado
              ]);
            });
          }
          nuevos++;
        }
      });
      return jsonOk({nuevos: nuevos});
    }

    return jsonOk({msg:'Accion desconocida: '+action});

  } catch(err) {
    return jsonOk({error: err.toString()});
  }
}

function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || '';

    if (action === 'getPendientes') {
      const sheet = getSheet('Anomalias');
      const datos = sheet.getDataRange().getValues();
      const pendientes = [];
      for (let i = 1; i < datos.length; i++) {
        const estado = String(datos[i][16] || '').trim().toLowerCase();
        if (estado !== 'cerrada') {
          pendientes.push({
            id:            datos[i][0],
            fecha:         datos[i][1] ? String(datos[i][1]).split('T')[0] : '',
            hora:          datos[i][2] || '',
            operador:      datos[i][3] || '',
            equipo:        datos[i][4] || '',
            disponibilidad:datos[i][5] || '',
            sede:          datos[i][6] || '',
            lote:          datos[i][7] || '',
            area:          datos[i][8] || '',
            sistema:       datos[i][9] || '',
            subsistema:    datos[i][10]|| '',
            parte:         datos[i][11]|| '',
            modo_falla:    datos[i][12]|| '',
            prioridad:     datos[i][13]|| '',
            falla:         datos[i][14]|| '',
            observaciones: datos[i][15]|| '',
            estado:        datos[i][16]|| 'Abierta',
            mecanismo:     datos[i][17]|| '',
            tecnico:       datos[i][18]|| '',
            fecha_ini_rep: datos[i][19]|| '',
            fecha_fin_rep: datos[i][20]|| '',
            obs_taller:    datos[i][21]|| '',
            fecha_estado:  datos[i][22]|| '',
            hist_estados:  [],
          });
        }
      }
      return jsonOk({pendientes: pendientes});
    }

    // Guardar lote via GET (cuando viene data en parámetro URL)
    if (action === 'saveLote' && e.parameter.data) {
      const sheet   = getSheet('Anomalias');
      const registros = JSON.parse(decodeURIComponent(e.parameter.data));
      const existe  = getIdsExistentes(sheet);
      let nuevos = 0;
      registros.forEach(a => {
        if (!existe.has(String(a.id))) {
          sheet.appendRow([a.id,a.fecha||'',a.hora||'',a.operador||'',a.equipo||'',
            a.disponibilidad||'',a.sede||'',a.lote||'',a.area||'',
            a.sistema||'',a.subsistema||'',a.parte||'',a.modo_falla||'',
            a.prioridad||'',a.falla||'',a.observaciones||'',
            a.estado||'Abierta',a.mecanismo||'',a.tecnico||'',
            a.fecha_ini_rep||'',a.fecha_fin_rep||'',a.obs_taller||'',
            a.fecha_estado||'',a.cod_requisicion||'',new Date().toISOString()]);
          nuevos++;
        }
      });
      return jsonOk({nuevos: nuevos});
    }

    if (action === 'ping') return jsonOk({msg:'OK'});

    // ── Devolver listas dinámicas para el APP2
    if (action === 'getListas') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const listas = {};

      // Helper para leer una hoja
      const leerHoja = (nombre) => {
        const sh = ss.getSheetByName(nombre);
        return sh ? sh.getDataRange().getValues() : null;
      };

      // Operadores
      const opRows = leerHoja('Lista_operadores');
      if (opRows) {
        listas.operadores = opRows.slice(1)
          .filter(r => r[0] && String(r[2]||'').trim().toUpperCase() === 'ACTIVO')
          .map(r => ({codigo: String(parseInt(r[0])), nombre: String(r[1]).trim()}));
      }

      // Equipos
      const eqRows = leerHoja('Lista_equipos');
      if (eqRows) {
        // Encontrar índice de columna Tiene_Implemento (col 6 = índice 6)
        const eqHeader = eqRows[0] || [];
        const implColIdx = eqHeader.map(c => String(c||'').trim()).indexOf('Tiene_Implemento');
        listas.equipos = eqRows.slice(1)
          .filter(r => r[0])
          .map(r => {
            const obj = {codigo: String(r[0]).trim(), tipo: String(r[1]||'').trim(), descripcion: String(r[4]||'').trim()};
            if (implColIdx >= 0) {
              obj.tiene_implemento = String(r[implColIdx]||'').trim().toUpperCase() === 'SI';
            }
            return obj;
          });
      }

      // Implementos
      const implRows = leerHoja('Lista_implementos');
      if (implRows) {
        listas.implementos = implRows.slice(1)
          .filter(r => r[0])
          .map(r => ({codigo: String(r[0]).trim(), descripcion: String(r[1]||'').trim()}));
      }

      // Labores (con clasificacion y distribucion)
      const labRows = leerHoja('Codigo_labor');
      if (labRows) {
        listas.labores = labRows.slice(1)
          .filter(r => r[0])
          .map(r => ({
            codigo: String(parseInt(r[0])),
            descripcion: String(r[1]||'').trim().replace(/ /g,''),
            clasificacion: String(r[2]||'').trim(),
            distribucion: String(r[3]||'').trim(),
            area: String(r[4]||'').trim(),
            unidad: String(r[5]||'').trim()
          }));
      }

      // Lotes
      const loteRows = leerHoja('Lista_lotes');
      if (loteRows) {
        listas.lotes = [];
        listas.lotes_por_finca = {};
        loteRows.slice(1).forEach(r => {
          if (!r[0]) return;
          const loteStr = String(r[0]).trim();
          const finca   = String(r[4]||'').trim() || 'General';
          listas.lotes.push({
            lote: loteStr,
            area_neta: parseFloat(r[1]) || 0,
            area_ef:   parseFloat(r[2]) || 0,
            estado:    String(r[3]||'').trim(),
            finca:     finca,
            distrito:  String(r[5]||'').trim()
          });
          if (!listas.lotes_por_finca[finca]) listas.lotes_por_finca[finca] = [];
          listas.lotes_por_finca[finca].push(loteStr);
        });
        const fincasSort = Object.keys(listas.lotes_por_finca).filter(f => f !== 'General').sort();
        if (listas.lotes_por_finca['General']) fincasSort.push('General');
        listas.fincas = fincasSort;
      }

      // Unidades
      const unidRows = leerHoja('Unidad');
      listas.unidades = unidRows
        ? unidRows.slice(1).filter(r => r[0]).map(r => String(r[0]).trim())
        : ['UNIDAD','METROS','HAS','BULTOS','UNIDADES','LITROS','BINES','VIAJES','KILOS','TONELADAS'];

      // Areas
      const areaRows = leerHoja('Lista_areas');
      listas.areas = areaRows
        ? areaRows.slice(1).filter(r => r[0]).map(r => String(r[0]).trim())
        : ["Aplicaciones","CTIA","Control operaciones","Cosecha","Desarrollo de cultivo","Empaque","Preparación","Taller","semillero","siembra"];

      return jsonOk({listas: listas, timestamp: new Date().toISOString()});
    }

    // ── Devolver listas dinámicas para el APP1 (Anomalías)
    if (action === 'getListasApp1') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const L  = {};

      const leerHoja = (nombre) => {
        const sh = ss.getSheetByName(nombre);
        return sh ? sh.getDataRange().getValues() : null;
      };

      // Operadores (activos)
      const opRows = leerHoja('Lista_operadores');
      if (opRows) {
        L.operadores = opRows.slice(1)
          .filter(r => r[0] && String(r[2]||'').trim().toUpperCase() === 'ACTIVO')
          .map(r => ({codigo: String(parseInt(r[0])), nombre: String(r[1]).trim()}));
      }

      // Equipos (tipo != IMPLEMENTO)
      const eqRows = leerHoja('Lista_equipos');
      if (eqRows) {
        L.equipos = eqRows.slice(1)
          .filter(r => r[0] && String(r[3]||'').trim().toUpperCase() !== 'IMPLEMENTO')
          .map(r => ({codigo: String(r[0]).trim(), descripcion: String(r[1]||'').trim(), tipo: 'EQUIPO'}));
      }

      // Implementos (de Lista_implementos)
      const implRows = leerHoja('Lista_implementos');
      if (implRows) {
        L.implementos = implRows.slice(1)
          .filter(r => r[0])
          .map(r => ({codigo: String(r[0]).trim(), descripcion: String(r[1]||'').trim(), tipo: 'IMPLEMENTO'}));
      }

      // Lotes + Fincas
      const loteRows = leerHoja('Lista_lotes');
      if (loteRows) {
        L.lotes = [];
        L.lotes_por_finca = {};
        loteRows.slice(1).forEach(r => {
          if (!r[0]) return;
          const loteStr = String(r[0]).trim();
          const finca   = String(r[4]||'').trim() || 'General';
          L.lotes.push(loteStr);
          if (!L.lotes_por_finca[finca]) L.lotes_por_finca[finca] = [];
          L.lotes_por_finca[finca].push(loteStr);
        });
        L.sedes = Object.keys(L.lotes_por_finca).sort();
      }

      // Áreas (intentar 'Lista_areas' primero, luego 'Areas' como fallback)
      const areaRows = leerHoja('Lista_areas') || leerHoja('Areas');
      L.areas = areaRows
        ? areaRows.slice(1).filter(r => r[0]).map(r => String(r[0]).trim())
        : ['Aplicaciones','CTIA','Control operaciones','Cosecha','Desarrollo de cultivo','Empaque','Preparación','Taller','semillero','siembra'];

      // Técnicos
      const tecRows = leerHoja('Técnico');
      L.tecnicos = tecRows
        ? tecRows.slice(1).filter(r => r[0]).map(r => String(r[0]).trim())
        : [];

      // Prioridades
      const prioRows = leerHoja('Prioridad');
      L.prioridades = prioRows
        ? prioRows.slice(1).filter(r => r[0]).map(r => String(r[0]).trim())
        : ['Alta','Media','Baja'];

      // Disponibilidades
      const dispRows = leerHoja('Disponibilidad');
      L.disponibilidades = dispRows
        ? dispRows.slice(1).filter(r => r[0]).map(r => String(r[0]).trim())
        : ['Varada','Disponible'];

      // Fallas → partes, sistemas, jerarquía, inv_partes, modos, mecanismos
      const fallaRows = leerHoja('Fallas');
      if (fallaRows) {
        const partesSet    = new Set();
        const sistemasSet  = new Set();
        const modosSet     = new Set();
        const mecSet       = new Set();
        const jerarquia    = {};
        const inv_partes   = {};

        fallaRows.slice(1).forEach(r => {
          const sis  = String(r[0]||'').trim();
          const sub  = String(r[1]||'').trim();
          const prt  = String(r[2]||'').trim();
          const modo = String(r[3]||'').trim();
          const mec  = String(r[4]||'').trim();
          if (!sis || !prt) return;

          partesSet.add(prt);
          sistemasSet.add(sis);
          if (modo) modosSet.add(modo);
          if (mec)  mecSet.add(mec);

          // Jerarquía
          if (!jerarquia[sis]) jerarquia[sis] = {};
          if (sub && !jerarquia[sis][sub]) jerarquia[sis][sub] = [];
          const arr = sub ? jerarquia[sis][sub] : (jerarquia[sis][''] = jerarquia[sis][''] || []);
          if (!arr.includes(prt)) arr.push(prt);

          // inv_partes
          const key = prt.toLowerCase();
          if (!inv_partes[key]) inv_partes[key] = [];
          const existe = inv_partes[key].some(x => x.sistema === sis && x.subsistema === sub);
          if (!existe) inv_partes[key].push({sistema: sis, subsistema: sub});
        });

        L.partes      = [...partesSet].sort();
        L.sistemas    = [...sistemasSet].sort();
        L.modos       = [...modosSet].sort();
        L.mecanismos  = [...mecSet].sort();
        L.jerarquia   = jerarquia;
        L.inv_partes  = inv_partes;
      }

      return jsonOk({listas: L, timestamp: new Date().toISOString()});
    }

    return jsonOk({msg:'Accion desconocida'});

  } catch(err) {
    return jsonOk({error: err.toString()});
  }
}

// ── Helpers
function getIdsExistentes(sheet) {
  const datos = sheet.getDataRange().getValues();
  const ids = new Set();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0]) ids.add(String(datos[i][0]));
  }
  return ids;
}

function jsonOk(extra) {
  const obj = Object.assign({ok: true}, extra||{});
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Borrar TODOS los datos (conserva encabezados) ──
function borrarTodo() {
  const sheet = getSheet('Anomalias');
  const ultFila = sheet.getLastRow();
  if (ultFila > 1) {
    sheet.deleteRows(2, ultFila - 1);
  }
  Logger.log('Hoja Anomalias limpiada. Solo quedan los encabezados.');
}

// ── Ejecutar manualmente desde Apps Script para limpiar duplicados ──
function limpiarDuplicados() {
  const sheet  = getSheet('Anomalias');
  const datos  = sheet.getDataRange().getValues();
  const vistos = new Set();
  const borrar = [];

  for (let i = datos.length - 1; i >= 1; i--) {
    const id = String(datos[i][0]).trim();
    if (!id || vistos.has(id)) {
      borrar.push(i + 1); // fila en Sheets (1-indexed)
    } else {
      vistos.add(id);
    }
  }

  // Borrar de abajo para arriba para no desplazar índices
  borrar.sort((a,b) => b - a);
  borrar.forEach(r => sheet.deleteRow(r));

  Logger.log('Duplicados eliminados: ' + borrar.length);
  Logger.log('Registros únicos restantes: ' + (datos.length - 1 - borrar.length));
  return { eliminados: borrar.length };
}
