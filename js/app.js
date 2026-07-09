      // Native XLSX export using JSZip and minimal OpenXML. Generates a real .xlsx file that opens correctly in Excel, Numbers, LibreOffice y OpenOffice.
      // This function mirrors the structure of exportExcelLikeXls but creates the XLSX zip package manually using JSZip.
      async function exportExcelXlsx() {
        // Build all sessions to export (current active plus day history).
        // Prefer helper from window if exposed, fallback to local definition.
        const sessions = (window.buildSessionsForExport || buildSessionsForExport)();
        if (!sessions || !sessions.length) {
          // Use toast helper from window if available
          (window.showToast || showToast)('No hay tandas para exportar');
          return;
        }
        // Load JSZip if it's not available globally. The library is stored locally in js/lib/jszip.min.js.
        if (typeof JSZip === 'undefined') {
          await new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'js/lib/jszip.min.js';
            script.onload = () => resolve();
            script.onerror = () => reject(new Error('No se pudo cargar JSZip'));
            document.head.appendChild(script);
          });
        }
        // Helper to escape characters that are invalid in XML.
        const esc = (str) => String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
        // Convert 0-based column index to Excel column letters (e.g. 0 -> A, 26 -> AA).
        const colLetter = (n) => {
          let s = '';
          n++;
          while (n > 0) {
            let r = (n - 1) % 26;
            s = String.fromCharCode(65 + r) + s;
            n = Math.floor((n - 1) / 26);
          }
          return s;
        };
        // Build styles.xml with several fills and XF definitions. We redefine the style sheet to include
        // alignment settings and an extra yellow fill for consistency ratings. Styles:
        // 0: default (white) left aligned; 1: header (grey) centered; 2: best lap (green) centered;
        // 3: worst lap (red) centered; 4: average (blue) centered; 5: consistency yellow centered;
        // 6: default centered; 7: unused/grey125.
        const buildStylesXML = () => {
          return (
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
            `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n` +
            `  <fonts count="1">\n` +
            `    <font>\n` +
            `      <sz val="11"/>\n` +
            `      <color theme="1"/>\n` +
            `      <name val="Calibri"/>\n` +
            `      <family val="2"/>\n` +
            `    </font>\n` +
            `  </fonts>\n` +
            // Fills: 0 none, 1 gray125 (unused), 2 header grey, 3 green, 4 red, 5 blue, 6 yellow
            `  <fills count="7">\n` +
            `    <fill><patternFill patternType="none"/></fill>\n` +
            `    <fill><patternFill patternType="gray125"/></fill>\n` +
            `    <fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>\n` +
            `    <fill><patternFill patternType="solid"><fgColor rgb="FFDFF0D8"/><bgColor indexed="64"/></patternFill></fill>\n` +
            `    <fill><patternFill patternType="solid"><fgColor rgb="FFF2DEDE"/><bgColor indexed="64"/></patternFill></fill>\n` +
            `    <fill><patternFill patternType="solid"><fgColor rgb="FFD9EDF7"/><bgColor indexed="64"/></patternFill></fill>\n` +
            `    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF4CE"/><bgColor indexed="64"/></patternFill></fill>\n` +
            `  </fills>\n` +
            `  <borders count="1">\n` +
            `    <border>\n` +
            `      <left style="thin"><color indexed="64"/></left>\n` +
            `      <right style="thin"><color indexed="64"/></right>\n` +
            `      <top style="thin"><color indexed="64"/></top>\n` +
            `      <bottom style="thin"><color indexed="64"/></bottom>\n` +
            `      <diagonal/>\n` +
            `    </border>\n` +
            `  </borders>\n` +
            `  <cellStyleXfs count="1">\n` +
            `    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>\n` +
            `  </cellStyleXfs>\n` +
            // Define 8 cell formats with alignment and fill references. wrapText=1 to allow multiline cells.
            `  <cellXfs count="8">\n` +
            // 0: default left aligned
            `    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>\n` +
            // 1: header grey, centered
            `    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>\n` +
            // 2: best lap green, centered
            `    <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>\n` +
            // 3: worst lap red, centered
            `    <xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>\n` +
            // 4: average blue, centered
            `    <xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>\n` +
            // 5: consistency yellow, centered
            `    <xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>\n` +
            // 6: default centered (no fill)
            `    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>\n` +
            // 7: unused grey125 (for completeness) left aligned
            `    <xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>\n` +
            `  </cellXfs>\n` +
            `  <cellStyles count="1">\n` +
            `    <cellStyle name="Normal" xfId="0" builtinId="0"/>\n` +
            `  </cellStyles>\n` +
            `</styleSheet>`
          );
        };
        // Dictionary to hold file contents before zipping
        const files = {};
        const sheetDataParts = [];
        // Create sheet for each session
        sessions.forEach((session, sIdx) => {
          const sheetIndex = sIdx + 1;
          const sheetName = `Tanda ${sheetIndex}`;
          // Build rows. Maintain column widths for auto-sizing.
          let rowNum = 1;
          let rowsXML = '';
          const colWidths = [];
          const addRow = (cells) => {
            // Update column widths based on cell contents
            cells.forEach((cell, cIdx) => {
              const val = cell && cell.v ? String(cell.v) : '';
              const lines = val.split('\n');
              const maxLen = lines.reduce((m, l) => Math.max(m, l.length), 0);
              if (colWidths[cIdx] === undefined || maxLen > colWidths[cIdx]) {
                colWidths[cIdx] = maxLen;
              }
            });
            // Build XML for this row
            let rxml = `<row r="${rowNum}">`;
            cells.forEach((cell, cIdx) => {
              if (!cell) {
                rxml += `<c r="${colLetter(cIdx)}${rowNum}"/>`;
              } else {
                const { v, s } = cell;
                const ref = `${colLetter(cIdx)}${rowNum}`;
                if (s !== undefined) {
                  rxml += `<c r="${ref}" t="inlineStr" s="${s}"><is><t>${esc(v)}</t></is></c>`;
                } else {
                  rxml += `<c r="${ref}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`;
                }
              }
            });
            rxml += '</row>';
            rowsXML += rxml;
            rowNum++;
          };
          // Determine start and end times
          let startEpoch = Infinity;
          let endEpoch = -Infinity;
          (session.riders || []).forEach(r => {
            (r.events || []).filter(e => e.kind === 'LAP').forEach(ev => {
              const se = Number(ev.stampEpoch);
              if (se < startEpoch) startEpoch = se;
              if (se > endEpoch) endEpoch = se;
            });
          });
          const sessionDate = session.dateISO ? new Date(session.dateISO) : (startEpoch !== Infinity ? new Date(startEpoch) : new Date());
          const fecha = sessionDate.toLocaleDateString(undefined, { year: 'numeric', month: '2-digit', day: '2-digit' });
          const horaInicio = startEpoch !== Infinity ? new Date(startEpoch).toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' }) : '';
          const horaFin = (endEpoch > startEpoch) ? new Date(endEpoch).toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' }) : '';
          // Construct header information. Only the following rows remain:
          // ENDUDO PRO, Fecha, Hora de inicio, Hora de finalización, Nombre de la tanda.
          addRow([{ v: 'ENDURO PRO', s: 1 }, null]);
          addRow([{ v: 'Fecha', s: 1 }, { v: fecha }]);
          addRow([{ v: 'Hora de inicio', s: 1 }, { v: horaInicio }]);
          addRow([{ v: 'Hora de finalización', s: 1 }, { v: horaFin }]);
          addRow([{ v: 'Nombre de la tanda', s: 1 }, { v: session.name || sheetName }]);
          // Empty row before table for spacing
          addRow([]);
          // Prepare rider lap data.  Include the full events list so that
          // we can derive sector times for each lap when exporting.
          const ridersData = (session.riders || []).map(r => {
            const laps = (r.events || []).filter(e => e.kind === 'LAP');
            return { name: r.name || '', laps, events: r.events || [] };
          });
          const maxLaps = Math.max(0, ...ridersData.map(r => r.laps.length));
          // Table header row. Include Piloto, laps, Promedio, Total de Tanda, Consistencia and Observación.
          const headerCells = [];
          headerCells.push({ v: 'Piloto', s: 1 });
          for (let i = 1; i <= maxLaps; i++) {
            headerCells.push({ v: 'V' + i, s: 1 });
          }
          headerCells.push({ v: 'Promedio', s: 1 });
          headerCells.push({ v: 'Total de Tanda', s: 1 });
          headerCells.push({ v: 'Consistencia', s: 1 });
          headerCells.push({ v: 'Observación', s: 1 });
          // Capture the row index of the header row for freeze panes.
          const freezeRow = rowNum;
          addRow(headerCells);
          // Determine the number of sectors recorded for this session.  A value
          // greater than zero indicates that partials are enabled and each lap
          // should display its sector breakdown vertically.  The snapshot
          // preserves this as `session.sectors` when the session was saved.
          const sectorsCount = Number(session.sectors || 0);

          // Data rows
          ridersData.forEach(r => {
            const lapTimesMs = r.laps.map(l => l.lapNetMs);
            let best = null;
            let worst = null;
            let avg = null;
            if (lapTimesMs.length) {
              best = Math.min(...lapTimesMs);
              worst = Math.max(...lapTimesMs);
              avg = lapTimesMs.reduce((a, b) => a + b, 0) / lapTimesMs.length;
            }
            const rowCells = [];
            // Pilot name remains left aligned (style 0 by default)
            rowCells.push({ v: r.name });
            // Lap cells
            for (let i = 0; i < maxLaps; i++) {
              const lap = r.laps[i];
              if (!lap) {
                // No lap recorded for this position; leave the cell blank
                rowCells.push(null);
              } else {
                const ms = lap.lapNetMs;
                // Format time as mm:ss (no hundredths)
                const timeStr = (() => {
                  const m = Math.floor(ms / 60000);
                  const s = Math.floor((ms % 60000) / 1000);
                  return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
                })();
                let pctLine = '';
                let styleIndex;
                if (avg) {
                  const ratio = (ms - avg) / avg;
                  const pct = Math.round(Math.abs(ratio) * 100);
                  const dir = ratio > 0 ? '↑' : (ratio < 0 ? '↓' : '');
                  pctLine = dir ? `${dir} ${pct}%` : `${pct}%`;
                  if (ms === best) styleIndex = 2; // green
                  else if (ms === worst) styleIndex = 3; // red
                  else if (Math.abs(ratio) < 0.01) styleIndex = 4; // blue for close to average
                  else styleIndex = 6; // default centered
                } else {
                  styleIndex = 6;
                }
                // If the session has sectors enabled, build a multi-line representation
                // for the lap that includes lap number, lap time, deviation from
                // the average and each sector time on its own line.  Otherwise
                // fall back to the original two-line layout (time + percentage).
                if (sectorsCount > 0) {
                  const lines = [];
                  // Lap number (use stored lapNo if available, otherwise index+1)
                  const lapNo = typeof lap.lapNo === 'number' ? lap.lapNo : (i + 1);
                  lines.push('V' + lapNo);
                  lines.push(timeStr);
                  lines.push(pctLine);
                  // Build sector times by looking back through the rider's events.
                  const events = r.events;
                  // Find the position of this lap event in the events array
                  const idxInEvents = events.indexOf(lap);
                  if (idxInEvents > -1) {
                    const sectorEvents = [];
                    // Walk backwards until we encounter the previous lap or collect all sectors
                    for (let j = idxInEvents - 1; j >= 0 && sectorEvents.length < sectorsCount; j--) {
                      const ev = events[j];
                      if (ev.kind === 'LAP') break;
                      if (ev.kind === 'SECTOR') {
                        sectorEvents.push(ev);
                      }
                    }
                    // Reverse to restore chronological order
                    sectorEvents.reverse();
                    // For each sector event, add a line "S# mm:ss"
                    sectorEvents.forEach(se => {
                      const secMs = se.deltaNetMs;
                      const m2 = Math.floor(secMs / 60000);
                      const s2 = Math.floor((secMs % 60000) / 1000);
                      const secStr = `${String(m2).padStart(2, '0')}:${String(s2).padStart(2, '0')}`;
                      // Label with sector number as stored in event.sectorNo
                      const label = typeof se.sectorNo === 'number' ? se.sectorNo : sectorEvents.indexOf(se) + 1;
                      lines.push(`S${label} ${secStr}`);
                    });
                  }
                  rowCells.push({ v: lines.join('\n'), s: styleIndex });
                } else {
                  // Default: time on first line and deviation on second line
                  rowCells.push({ v: timeStr + '\n' + pctLine, s: styleIndex });
                }
              }
            }
            if (avg !== null) {
              // Average time mm:ss
              const avgStr = (() => {
                const m = Math.floor(avg / 60000);
                const s = Math.floor((avg % 60000) / 1000);
                return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
              })();
              rowCells.push({ v: avgStr, s: 4 });
              // Total time mm:ss (sum of laps)
              const totalMs = lapTimesMs.reduce((a, b) => a + b, 0);
              const totalStr = (() => {
                const m = Math.floor(totalMs / 60000);
                const s = Math.floor((totalMs % 60000) / 1000);
                return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
              })();
              rowCells.push({ v: totalStr, s: 6 });
              // Consistency: coefficient of variation (std dev / mean) as percentage (integer)
              let cvPct = 0;
              if (lapTimesMs.length > 1) {
                const mean = avg;
                const variance = lapTimesMs.reduce((acc, b) => acc + Math.pow(b - mean, 2), 0) / lapTimesMs.length;
                const stdDev = Math.sqrt(variance);
                cvPct = Math.round((stdDev / mean) * 100);
              }
              let cvStyle;
              if (cvPct <= 3) cvStyle = 2; // green
              else if (cvPct <= 6) cvStyle = 5; // yellow
              else cvStyle = 3; // red
              rowCells.push({ v: `${cvPct}%`, s: cvStyle });
              // Observación
              let obs = '';
              if (lapTimesMs.length < 2) {
                obs = 'Datos insuficientes';
              } else {
                const half = Math.floor(lapTimesMs.length / 2);
                const firstHalf = lapTimesMs.slice(0, half);
                const secondHalf = lapTimesMs.slice(half);
                const avgFirst = firstHalf.reduce((a, b) => a + b, 0) / firstHalf.length;
                const avgSecond = secondHalf.reduce((a, b) => a + b, 0) / secondHalf.length;
                const diffRatio = (avgSecond - avgFirst) / avgFirst;
                const slowLaps = lapTimesMs.filter(t => (t - best) / best > 0.15).length;
                if (slowLaps >= 1) obs = 'Hubo una vuelta claramente fuera del promedio.';
                else if (diffRatio > 0.05) obs = 'Caída de rendimiento al final.';
                else if (diffRatio < -0.05) obs = 'Mejora hacia el final.';
                else if (cvPct <= 3) obs = 'Muy constante.';
                else if (cvPct <= 6) obs = 'Buen ritmo.';
                else if (cvPct <= 10) obs = 'Ritmo irregular.';
                else obs = 'Ritmo muy irregular.';
              }
              rowCells.push({ v: obs });
            } else {
              // No laps: empty average, total, consistency and observation
              rowCells.push(null);
              rowCells.push(null);
              rowCells.push(null);
              rowCells.push({ v: 'Sin vueltas' });
            }
            addRow(rowCells);
          });
          // Build column definitions based on measured widths. Add some padding for readability.
          const colsXmlParts = colWidths.map((len, idx) => {
            // Minimum width of 8 characters, plus small padding
            const width = Math.max(8, len + 2);
            return `<col min="${idx + 1}" max="${idx + 1}" width="${width}" customWidth="1"/>`;
          }).join('');
          const colsXml = `<cols>${colsXmlParts}</cols>`;
          // Build sheet view with frozen panes (freeze header row and first column)
          // freezeRow variable captured before adding header row; freeze first column (A)
          const xSplit = 1;
          const ySplit = freezeRow;
          const topLeftCell = `${colLetter(xSplit)}${ySplit + 1}`;
          const sheetViews = `<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="${xSplit}" ySplit="${ySplit}" topLeftCell="${topLeftCell}" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="${topLeftCell}" sqref="${topLeftCell}"/></sheetView></sheetViews>`;
          // Assemble final sheet XML with sheetViews, sheetFormatPr, cols and sheetData
          const sheetXml =
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
            `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n` +
            `  ${sheetViews}\n` +
            `  <sheetFormatPr defaultRowHeight="15"/>\n` +
            `  ${colsXml}\n` +
            `  <sheetData>${rowsXML}</sheetData>\n` +
            `</worksheet>`;
          const sheetPath = `xl/worksheets/sheet${sheetIndex}.xml`;
          files[sheetPath] = sheetXml;
          sheetDataParts.push({ name: sheetName, path: sheetPath });
        });
        // Build workbook.xml
        const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
          `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n` +
          `  <sheets>\n` +
          sheetDataParts.map((sh, idx) => `    <sheet name="${esc(sh.name)}" sheetId="${idx + 1}" r:id="rId${idx + 1}"/>`).join('\n') + `\n` +
          `  </sheets>\n` +
          `</workbook>`;
        files['xl/workbook.xml'] = workbookXml;
        // Build workbook relationships
        const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
          `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n` +
          sheetDataParts.map((sh, idx) => `  <Relationship Id="rId${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${idx + 1}.xml"/>`).join('\n') + `\n` +
          `  <Relationship Id="rId${sheetDataParts.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n` +
          `</Relationships>`;
        files['xl/_rels/workbook.xml.rels'] = workbookRels;
        // Build [Content_Types].xml
        const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
          `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n` +
          `  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n` +
          `  <Default Extension="xml" ContentType="application/xml"/>\n` +
          `  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\n` +
          `  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>\n` +
          sheetDataParts.map((sh, idx) => `  <Override PartName="/xl/worksheets/sheet${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('\n') + `\n` +
          `</Types>`;
        files['[Content_Types].xml'] = contentTypes;
        // Root relationships
        files['_rels/.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
          `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n` +
          `  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n` +
          `</Relationships>`;
        // Add styles.xml
        files['xl/styles.xml'] = buildStylesXML();
        // Use JSZip to assemble the zip archive
        const zip = new JSZip();
        Object.keys(files).forEach(path => {
          zip.file(path, files[path]);
        });
        const xlsxBuffer = await zip.generateAsync({ type: 'arraybuffer' });
        const blob = new Blob([xlsxBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'jornada_enduro_pro_' + new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-') + '.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
        (window.showToast || showToast)('Informe Excel nativo exportado');
      }
(() => {
      const $ = id => document.getElementById(id), STORAGE_KEY = 'cronometro_enduro_v5', DAY_KEY = 'cronometro_enduro_v5_day_history'; const DEFAULT_NAMES = ['Esteban', 'Guille', 'Ariel', 'Richard', 'Gera', 'Fran', 'Ale', 'Franco']; const PALETTE = ['#2a4cff', '#00c2ff', '#00d084', '#ffb000', '#ff5c7a', '#b56bff', '#7bdcff', '#a7ff4f', '#ff7a00', '#00ffa8', '#ff3df2', '#3dff57', '#ffd500', '#6bb3ff', '#ff6bb3', '#a56bff']; const elTime = $('time'), elSub = $('sub'), btnStartPause = $('startPause'), btnReset = $('reset'), btnExport = $('export'), btnClearAll = $('clearAll'), btnClearDay = $('clearDay'), btnFinishSession = $('finishSession'), btnToggleButtons = $('toggleButtons'), elPrepSec = $('prepSec'), elGap = $('gap'), elDblMs = $('dblMs'), elSectors = $('sectors'), elTrackName = $('trackName'), elSessionName = $('sessionName'), elSessionNotes = $('sessionNotes'), elSetupNotes = $('setupNotes'), ridersBtns = $('ridersBtns'), ridersList = $('ridersList'), historyList = $('historyList'), tables = $('tables'), btnAddRider = $('addRider'), btnResetNames = $('resetNames'), toast = $('toast'), launchOverlay = $('launchOverlay'), launchTitle = $('launchTitle'), launchCount = $('launchCount'), launchPilot = $('launchPilot'), launchHint = $('launchHint'); let running = false, startPerf = 0, elapsedMs = 0, rafId = null, launchMode = false, riders = [], dayHistory = [], nextIndex = 0, confirmMap = {}, raceStartElapsedMs = null, raceBaseNetById = {}; function uid() { return Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(2, 8) } function showToast(m) { toast.textContent = m; toast.classList.add('show'); setTimeout(() => toast.classList.remove('show'), 1300) } function format(ms) { ms = Math.max(0, Math.floor(ms)); let ts = Math.floor(ms / 1000), m = Math.floor(ts / 60), s = ts % 60, c = Math.floor((ms % 1000) / 10); return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}.${String(c).padStart(2, '0')}` } function formatQuick(ms) { ms = Math.max(0, Math.floor(ms)); let ts = Math.floor(ms / 1000), m = Math.floor(ts / 60), s = ts % 60; return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}` } function nowElapsed() { return running ? elapsedMs + (performance.now() - startPerf) : elapsedMs } function getPrepSec() { return Math.max(0, Number(elPrepSec.value || 20)) } function getGapMs() { return Math.max(0, Number(elGap.value || 0)) * 1000 } function getDblMs() { return Math.max(0, Number(elDblMs.value || 0)) } function getSectors() { return Math.max(0, Number(elSectors.value || 0)) } function offsetMsFor(i) { return i * getGapMs() } function totalNetMs(i, t) { return Math.max(0, t - offsetMsFor(i)) } function headerText() { if (launchMode) return 'Secuencia de largada…'; if (!running) return elapsedMs > 0 ? 'Pausado' : 'Listo'; return getSectors() ? `Corriendo… (parciales ${getSectors()})` : 'Corriendo… (vueltas)' } function lapsOf(r) { return r.events.filter(e => e.kind === 'LAP') } function countLaps(r) { return lapsOf(r).length } function lastLap(r) { let l = lapsOf(r); return l.length ? l[l.length - 1] : null } function ensureDefaults() { if (riders.length) return; riders = DEFAULT_NAMES.map((name, i) => ({ id: uid(), name, color: PALETTE[i % PALETTE.length], events: [], sectorIdx: 0, lastTapEpoch: 0, redo: [] })) } function renderHeader() { elTime.textContent = format(nowElapsed()); elSub.textContent = headerText(); btnReset.disabled = !running && elapsedMs === 0 && riders.every(r => r.events.length === 0); btnExport.disabled = dayHistory.length === 0 && riders.every(r => countLaps(r) === 0); btnClearAll.disabled = riders.every(r => r.events.length === 0); btnFinishSession.disabled = riders.every(r => r.events.length === 0); btnToggleButtons.disabled = false } function renderButtons() {
        ridersBtns.innerHTML = '';
        let sec = getSectors();

        riders.forEach((r, idx) => {
          let b = document.createElement('button');
          b.className = 'riderBtn';
          if (idx === nextIndex) b.classList.add('next');
          if (confirmMap[r.id]) b.classList.add('confirm');
          b.style.borderColor = r.color;
          b.disabled = !running || launchMode;

          let lapsN = countLaps(r);
          let ll = lastLap(r);
          let last = ll ? formatQuick(ll.lapNetMs) : '--:--';
          let nextLapNo = lapsN + 1;

          if (confirmMap[r.id]) {
            b.innerHTML = `<div class="name">${r.name}</div>
        <div class="line1">${confirmMap[r.id].label}</div>
        <div style="font-size:42px;font-weight:1000;font-variant-numeric:tabular-nums;line-height:1">${confirmMap[r.id].time}</div>`;
          } else {
            b.innerHTML = `<div class="name">${r.name}</div>
        <div class="line1">Vuelta ${nextLapNo}</div>
        <div class="line2"><span class="label">Última</span><span class="val">${last}</span></div>
        <div class="line2"><span class="label">Parcial</span><span class="val">${sec ? `P${r.sectorIdx + 1}/${sec}` : 'OFF'}</span></div>`;
          }
          b.onclick = () => mark(idx);
          ridersBtns.appendChild(b);
        });
      } function renderConfigList() { ridersList.innerHTML = ''; riders.forEach((r, idx) => { let row = document.createElement('div'); row.className = 'rline'; let nameWrap = document.createElement('div'); nameWrap.className = 'nameWrap'; let dot = document.createElement('div'); dot.className = 'colorDot'; dot.style.borderColor = r.color; dot.style.background = r.color + '22'; let inp = document.createElement('input'); inp.value = r.name; inp.oninput = () => { r.name = (inp.value || '').trim() || `Piloto ${idx + 1}`; renderButtons(); renderTables(); save() }; nameWrap.append(dot, inp); let up = document.createElement('button'); up.textContent = '↑'; up.disabled = idx === 0; up.onclick = e => { e.preventDefault(); move(idx, -1) }; let down = document.createElement('button'); down.textContent = '↓'; down.disabled = idx === riders.length - 1; down.onclick = e => { e.preventDefault(); move(idx, 1) }; let del = document.createElement('button'); del.textContent = 'Borrar'; del.className = 'warn'; del.disabled = riders.length <= 1; del.onclick = e => { e.preventDefault(); removeRider(idx) }; row.append(nameWrap, up, down, del); ridersList.appendChild(row) }); let hint = document.createElement('div'); hint.innerHTML = '<small style="opacity:.75">Orden = orden de largada. Podés reordenar con ↑ / ↓.</small>'; hint.style.marginTop = '6px'; ridersList.appendChild(hint) } function renderHistory() {
        historyList.innerHTML = '';
        if (!dayHistory.length) {
          historyList.innerHTML = '<div style="opacity:.7">Sin tandas guardadas todavía.</div>';
          return;
        }

        dayHistory.slice().reverse().forEach((s, rev) => {
          let idx = dayHistory.length - 1 - rev;
          let div = document.createElement('div');
          div.className = 'historyItem';
          div.style.display = 'block';

          let inner = `<div style="display:grid;grid-template-columns:1fr auto;gap:8px;align-items:center">
      <div><b>${s.name || ('Tanda ' + (idx + 1))}</b><br><small>${s.dateText} · ${s.track || 'Sin pista'}</small></div>
      <button data-i="${idx}" class="warn" style="padding:8px 10px;font-size:12px">Borrar</button>
    </div>`;

          inner += `<div style="margin-top:8px;display:grid;gap:6px">`;
          s.riders.forEach(r => {
            let laps = r.events.filter(e => e.kind === 'LAP');
            if (!laps.length) return;
            let best = laps.reduce((b, c) => c.lapNetMs < b.lapNetMs ? c : b, laps[0]);
            inner += `<div style="background:#0e1526;border:1px solid #1f2a44;border-radius:10px;padding:8px">
        <div style="font-weight:950;margin-bottom:4px">${r.name}</div>
        <div style="font-variant-numeric:tabular-nums;display:flex;gap:8px;flex-wrap:wrap">`;
            laps.forEach(l => {
              let star = l.lapNetMs === best.lapNetMs ? '⭐' : '';
              inner += `<span>${star}${formatQuick(l.lapNetMs)}</span>`;
            });
            inner += `</div></div>`;
          });
          inner += `</div>`;
          div.innerHTML = inner;
          historyList.appendChild(div);
        });

        historyList.querySelectorAll('button[data-i]').forEach(btn => btn.onclick = () => {
          let i = Number(btn.dataset.i);
          if (confirm('¿Borrar esta tanda de la jornada?')) {
            dayHistory.splice(i, 1);
            saveDayHistory();
            renderHistory();
            renderHeader();
          }
        });
      } function renderTables() {
        tables.innerHTML = '';
        riders.forEach((r, idx) => {
          let card = document.createElement('div');
          card.className = 'tableCard';
          card.innerHTML = `
            <h3>${r.name} <span style="opacity:.7;font-weight:650">(offset ${Math.round(offsetMsFor(idx) / 1000)}s)</span></h3>
            <div style="display:flex;gap:8px;margin-bottom:8px;flex-wrap:wrap">
              <button data-undo="${idx}" ${r.events.length ? '' : 'disabled'}>Deshacer</button>
              <button data-redo="${idx}" ${r.redo.length ? '' : 'disabled'}>Rehacer</button>
              <button data-clear="${idx}" ${r.events.length ? '' : 'disabled'}>Vaciar</button>
            </div>
            <div class="scroll">
              <table>
                <thead>
                  <tr><th>#</th><th>Tipo</th><th>Δ neto</th><th>Total neto</th><th></th></tr>
                </thead>
                <tbody>
                  ${r.events.length ? r.events.map((e, evIdx) => `
                    <tr>
                      <td>${e.kind === 'LAP' ? e.lapNo : ''}</td>
                      <td>${e.kind === 'LAP' ? 'VUELTA' : 'S' + e.sectorNo}</td>
                      <td>${format(e.deltaNetMs)}</td>
                      <td>${format(e.totalNetMs)}</td>
                      <td><button data-del-event="${idx}:${evIdx}" class="warn" style="padding:5px 8px;font-size:11px">X</button></td>
                    </tr>`).join('') : '<tr><td colspan="5" style="opacity:.7">Sin registros todavía</td></tr>'}
                </tbody>
              </table>
            </div>`;
          tables.appendChild(card);
        });
        tables.querySelectorAll('button[data-undo]').forEach(b => b.onclick = () => undo(Number(b.dataset.undo)));
        tables.querySelectorAll('button[data-redo]').forEach(b => b.onclick = () => redo(Number(b.dataset.redo)));
        tables.querySelectorAll('button[data-clear]').forEach(b => b.onclick = () => clearRider(Number(b.dataset.clear)));
        tables.querySelectorAll('button[data-del-event]').forEach(b => b.onclick = () => {
          let parts = b.dataset.delEvent.split(':');
          deleteEvent(Number(parts[0]), Number(parts[1]));
        });
      }

      function deleteEvent(riderIdx, eventIdx) {
        let r = riders[riderIdx];
        if (!r || !r.events[eventIdx]) return;
        if (!confirm('¿Eliminar este registro del análisis/exportación?')) return;
        r.redo.push(r.events[eventIdx]);
        r.events.splice(eventIdx, 1);
        recomputeLaps(r);
        recomputeSectorIdx(r);
        renderAll();
        save();
      }

      function recomputeLaps(r) {
        let lapNo = 1;
        let totalNet = 0;

        r.events.forEach(e => {
          if (e.kind === 'LAP') {
            // Si se elimina una vuelta anterior, NO hay que recalcular la vuelta
            // usando el total acumulado viejo. La duración real de cada vuelta ya
            // está guardada en lapNetMs/deltaNetMs y debe conservarse.
            const lapDuration = Math.max(0, Number(e.lapNetMs ?? e.deltaNetMs ?? 0));
            totalNet += lapDuration;

            e.lapNo = lapNo++;
            e.lapNetMs = lapDuration;
            e.deltaNetMs = lapDuration;
            e.totalNetMs = totalNet;
            if (typeof e.riderIndex === 'number') e.totalMs = totalNet + offsetMsFor(e.riderIndex);
          }
        });
      } function renderAll() { renderHeader(); renderButtons(); renderConfigList(); renderTables(); renderHistory() } function tick() { renderHeader(); if (running) rafId = requestAnimationFrame(tick) } function saveDayHistory() { localStorage.setItem(DAY_KEY, JSON.stringify(dayHistory)) } function loadDayHistory() { try { dayHistory = JSON.parse(localStorage.getItem(DAY_KEY) || '[]') } catch { dayHistory = [] } } function save() { let s = { running, elapsedMs, startedAtEpoch: running ? Date.now() : null, prepSec: Number(elPrepSec.value || 20), gapSec: Number(elGap.value || 0), dblMs: Number(elDblMs.value || 0), sectors: Number(elSectors.value || 0), trackName: elTrackName.value || '', sessionName: elSessionName.value || '', sessionNotes: elSessionNotes.value || '', setupNotes: elSetupNotes.value || '', riders, nextIndex }; localStorage.setItem(STORAGE_KEY, JSON.stringify(s)) } function load() { let raw = localStorage.getItem(STORAGE_KEY); if (!raw) { ensureDefaults(); return } try { let s = JSON.parse(raw); if (typeof s.prepSec === 'number') elPrepSec.value = String(s.prepSec); if (typeof s.gapSec === 'number') elGap.value = String(s.gapSec); if (typeof s.dblMs === 'number') elDblMs.value = String(s.dblMs); if (typeof s.sectors === 'number') elSectors.value = String(s.sectors); elTrackName.value = s.trackName || ''; elSessionName.value = s.sessionName || ''; elSessionNotes.value = s.sessionNotes || ''; elSetupNotes.value = s.setupNotes || ''; nextIndex = typeof s.nextIndex === 'number' ? s.nextIndex : 0; if (Array.isArray(s.riders) && s.riders.length) riders = s.riders.map((r, i) => ({ id: r.id || uid(), name: r.name || DEFAULT_NAMES[i] || `Piloto ${i + 1}`, color: r.color || PALETTE[i % PALETTE.length], events: Array.isArray(r.events) ? r.events : [], sectorIdx: Number(r.sectorIdx || 0), lastTapEpoch: Number(r.lastTapEpoch || 0), redo: Array.isArray(r.redo) ? r.redo : [] })); else ensureDefaults(); elapsedMs = Number(s.elapsedMs || 0); if (s.running && s.startedAtEpoch) elapsedMs += Math.max(0, Date.now() - Number(s.startedAtEpoch)); running = false; nextIndex = Math.min(Math.max(0, nextIndex), riders.length - 1) } catch { ensureDefaults(); nextIndex = 0 } } function startTimer() { running = true; startPerf = performance.now(); btnStartPause.textContent = 'Pausar'; if (!rafId) rafId = requestAnimationFrame(tick); if (nextIndex < 0 || nextIndex >= riders.length) nextIndex = 0; renderAll(); save() } function getLastNetForRider(r) { if (!r || !Array.isArray(r.events) || !r.events.length) return 0; let lastLap = [...r.events].reverse().find(e => e.kind === 'LAP'); if (lastLap && typeof lastLap.totalNetMs === 'number') return lastLap.totalNetMs; let lastEvent = r.events[r.events.length - 1]; return typeof lastEvent.totalNetMs === 'number' ? lastEvent.totalNetMs : 0 } function buildRaceBaseNetById() { let base = {}; riders.forEach(r => { base[r.id] = getLastNetForRider(r) }); return base } function start() { raceBaseNetById = buildRaceBaseNetById(); raceStartElapsedMs = null; startTimer(); runLaunchSequence() } function pause() { elapsedMs = nowElapsed(); running = false; btnStartPause.textContent = 'Iniciar'; if (rafId) { cancelAnimationFrame(rafId); rafId = null } renderAll(); save() } function sleep(ms) { return new Promise(r => setTimeout(r, ms)) } let audioCtx = null;

      function beep(freq = 900, duration = 120) {
        try {
          audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)();
          const osc = audioCtx.createOscillator();
          const gain = audioCtx.createGain();

          osc.type = 'sine';
          osc.frequency.value = freq;
          gain.gain.value = 0.20;

          osc.connect(gain);
          gain.connect(audioCtx.destination);

          osc.start();
          osc.stop(audioCtx.currentTime + duration / 1000);
        } catch (e) { }
      }

      async function startBeep() {
        beep(950, 90);
        await sleep(120);
        beep(1100, 90);
        await sleep(120);
        beep(1300, 260);
      } async function runLaunchSequence() {
        if (launchMode) return;
        launchMode = true;
        renderButtons();

        try {
          launchOverlay.classList.add('show');
          launchOverlay.classList.remove('green');

          let prep = getPrepSec();
          launchTitle.textContent = 'PREPARADOS';
          launchPilot.textContent = '';
          launchHint.textContent = 'Cuenta inicial';

          for (let i = prep; i > 0; i--) {
            if (!running) break;
            launchOverlay.classList.remove('green');
            launchCount.textContent = String(i);

            if (i <= 5) beep(850, 120);

            await sleep(1000);
          }

          for (let i = 0; i < riders.length; i++) {
            if (!running) break;

            // Este instante es la largada real. El verde coincide con el 0.
            launchOverlay.classList.add('green');
            launchTitle.textContent = 'LARGA';
            launchPilot.textContent = riders[i].name;
            launchCount.textContent = '0';
            launchHint.textContent = `Piloto ${i + 1} de ${riders.length}`;
            startBeep();

            if (i === 0) {
              raceStartElapsedMs = nowElapsed();
            }

            const launchAt = performance.now();

            // Flash corto, sin sumar 2 segundos muertos al gap.
            await sleep(1000);

            if (i < riders.length - 1) {
              const target = launchAt + getGapMs();
              launchOverlay.classList.remove('green');
              launchTitle.textContent = 'PRÓXIMA LARGADA';
              launchPilot.textContent = riders[i + 1].name;
              launchHint.textContent = 'Gap de largada';

              let ultimoSegundoSonado = null;

              while (running) {
                let remain = Math.ceil((target - performance.now()) / 1000);
                if (remain <= 0) break;

                launchCount.textContent = String(remain);

                if (remain <= 5 && remain !== ultimoSegundoSonado) {
                  beep(850, 120);
                  ultimoSegundoSonado = remain;
                }

                await sleep(100);
              }

            }
          }


          if (running) {
            launchOverlay.classList.add('green');
            launchTitle.textContent = 'Aceleraaaa';
            launchPilot.textContent = 'Campeón';
            launchCount.textContent = '';
            launchHint.textContent = 'Tocá cada piloto cuando pase';
            await sleep(1000);
          }
        } finally {
          launchMode = false;
          launchOverlay.classList.remove('show', 'green');
          document.body.classList.add('buttonsOnly');
          renderAll();
          save();
        }
      } function resetTimeOnly() { running = false; launchMode = false; raceStartElapsedMs = null; raceBaseNetById = {}; launchOverlay.classList.remove('show', 'green'); elapsedMs = 0; startPerf = 0; if (rafId) { cancelAnimationFrame(rafId); rafId = null } btnStartPause.textContent = 'Iniciar'; renderAll(); save() } function clearAll() { raceStartElapsedMs = null; raceBaseNetById = {}; riders.forEach(r => { r.events = []; r.redo = []; r.sectorIdx = 0; r.lastTapEpoch = 0 }); nextIndex = 0; elapsedMs = 0; running = false; launchMode = false; if (rafId) { cancelAnimationFrame(rafId); rafId = null } btnStartPause.textContent = 'Iniciar'; renderAll(); save() } function snapshotCurrentSession() { return { id: uid(), name: elSessionName.value || `Tanda ${dayHistory.length + 1}`, track: elTrackName.value || '', notes: elSessionNotes.value || '', setupNotes: elSetupNotes.value || '', dateISO: new Date().toISOString(), dateText: new Date().toLocaleString(), prepSec: getPrepSec(), gapSec: Math.round(getGapMs() / 1000), sectors: getSectors(), riders: JSON.parse(JSON.stringify(riders)) } } function finishSession() { if (riders.every(r => r.events.length === 0)) { showToast('No hay tiempos para guardar'); return } dayHistory.push(snapshotCurrentSession()); saveDayHistory(); showToast('Tanda guardada ✔'); clearAll(); if (!elSessionName.value) elSessionName.value = `Tanda ${dayHistory.length + 1}`; renderAll(); save() } function recomputeSectorIdx(r) { let sec = getSectors(); if (sec <= 0) { r.sectorIdx = 0; return } let last = [...r.events].reverse().find(e => e.kind === 'SECTOR'); r.sectorIdx = last ? Number(last.sectorNo) % sec : 0 } function mark(idx) {
        let tEpoch = Date.now(), dbl = getDblMs(), r = riders[idx];
        if (dbl > 0 && (tEpoch - (r.lastTapEpoch || 0)) < dbl) {
          if (!confirm(`¿Confirmás el tap seguido para ${r.name}?`)) return;
        }
        r.lastTapEpoch = tEpoch;
        r.redo = [];

        let totalMs = nowElapsed();
        let netMs;

        if (raceStartElapsedMs !== null) {
          // When the timer is paused and started again, the new launch sequence
          // starts a new timing segment. Each rider must continue from their
          // own previous total, otherwise the first lap after resuming becomes 0.
          let baseNet = Number(raceBaseNetById[r.id] || 0);
          let segmentMs = totalMs - raceStartElapsedMs - offsetMsFor(idx);
          netMs = baseNet + Math.max(0, segmentMs);
          totalMs = netMs + offsetMsFor(idx);
        } else {
          netMs = totalNetMs(idx, totalMs);
        }

        let sec = getSectors();
        let lastNet = r.events.length ? r.events[r.events.length - 1].totalNetMs : 0;
        let deltaNet = netMs - lastNet;
        let label = '', timeMs = deltaNet, isBest = false, displayMs = 2000;

        if (sec > 0) {
          let sectorNo = r.sectorIdx + 1;
          r.events.push({ kind: 'SECTOR', riderIndex: idx, riderName: r.name, sectorNo, totalMs, totalNetMs: netMs, deltaNetMs: deltaNet, stampEpoch: tEpoch });
          label = `P${sectorNo}/${sec}`;
          r.sectorIdx++;

          if (r.sectorIdx >= sec) {
            let previousLaps = r.events.filter(x => x.kind === 'LAP');
            let lapNo = previousLaps.length + 1;
            let prev = previousLaps.length ? previousLaps[previousLaps.length - 1] : null;
            let prevTotal = prev ? prev.totalNetMs : 0;
            let lapNet = netMs - prevTotal;
            let prevBest = previousLaps.length ? Math.min(...previousLaps.map(x => x.lapNetMs)) : Infinity;
            isBest = lapNet < prevBest;
            r.events.push({ kind: 'LAP', riderIndex: idx, riderName: r.name, lapNo, totalMs, totalNetMs: netMs, deltaNetMs: lapNet, lapNetMs: lapNet, stampEpoch: tEpoch });
            r.sectorIdx = 0;
            label = `${isBest ? '⭐ ' : ''}Vuelta ${lapNo}`;
            timeMs = lapNet;
            displayMs = isBest ? 3000 : 2000;
          }
        } else {
          let previousLaps = r.events.filter(x => x.kind === 'LAP');
          let lapNo = previousLaps.length + 1;
          let prev = previousLaps.length ? previousLaps[previousLaps.length - 1] : null;
          let prevTotal = prev ? prev.totalNetMs : 0;
          let lapNet = netMs - prevTotal;
          let prevBest = previousLaps.length ? Math.min(...previousLaps.map(x => x.lapNetMs)) : Infinity;
          isBest = lapNet < prevBest;
          r.events.push({ kind: 'LAP', riderIndex: idx, riderName: r.name, lapNo, totalMs, totalNetMs: netMs, deltaNetMs: lapNet, lapNetMs: lapNet, stampEpoch: tEpoch });
          label = `${isBest ? '⭐ ' : ''}Vuelta ${lapNo}`;
          timeMs = lapNet;
          displayMs = isBest ? 3000 : 2000;
        }

        confirmMap[r.id] = { label, time: formatQuick(timeMs) };
        setTimeout(() => { delete confirmMap[r.id]; renderButtons() }, displayMs);

        nextIndex = (idx + 1) % riders.length;
        renderButtons();
        renderTables();
        save();
      } function undo(idx) { let r = riders[idx], last = r.events.pop(); if (!last) return; r.redo.push(last); recomputeSectorIdx(r); renderAll(); save() } function redo(idx) { let r = riders[idx], item = r.redo.pop(); if (!item) return; r.events.push(item); recomputeSectorIdx(r); renderAll(); save() } function clearRider(idx) { let r = riders[idx]; r.events = []; r.redo = []; r.sectorIdx = 0; r.lastTapEpoch = 0; renderAll(); save() } function addRider(name = '') { let i = riders.length; riders.push({ id: uid(), name: name || `Piloto ${i + 1}`, color: PALETTE[i % PALETTE.length], events: [], sectorIdx: 0, lastTapEpoch: 0, redo: [] }); renderAll(); save() } function removeRider(idx) { let r = riders[idx]; if (!confirm(`¿Borrar a ${r.name}?`)) return; riders.splice(idx, 1); nextIndex = 0; renderAll(); save() } function move(idx, dir) { let j = idx + dir; if (j < 0 || j >= riders.length) return;[riders[idx], riders[j]] = [riders[j], riders[idx]]; nextIndex = 0; renderAll(); save() } function restoreNames() { riders = DEFAULT_NAMES.map((name, i) => ({ id: uid(), name, color: PALETTE[i % PALETTE.length], events: [], sectorIdx: 0, lastTapEpoch: 0, redo: [] })); nextIndex = 0; renderAll(); save() } function buildSessionsForExport() { let s = dayHistory.slice(); if (!riders.every(r => r.events.length === 0)) s.push(snapshotCurrentSession()); return s } function exportExcelLikeXls() {
        // Build all sessions to export (current active plus day history).
        const sessions = buildSessionsForExport();
        if (!sessions.length) {
          showToast('No hay tandas para exportar');
          return;
        }
        // Helpers to format durations and dates
        const formatDuration = (ms) => {
          if (ms === null || ms === undefined || isNaN(Number(ms))) return '';
          ms = Math.max(0, Math.floor(Number(ms)));
          const totalSeconds = Math.floor(ms / 1000);
          const minutes = Math.floor(totalSeconds / 60);
          const seconds = totalSeconds % 60;
          const centis = Math.floor((ms % 1000) / 10);
          return String(minutes).padStart(2, '0') + ':' + String(seconds).padStart(2, '0') + '.' + String(centis).padStart(2, '0');
        };
        const formatDate = (d) => {
          return d.toLocaleDateString(undefined, { year: 'numeric', month: '2-digit', day: '2-digit' });
        };
        const formatTime = (d) => {
          return d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
        };
        // Build HTML for Excel export. We use HTML tables styled inline so that Excel
        // preserves colours, borders and layout when opening the .xls.
        let html = '<html><head><meta charset="UTF-8"><style>';
        html += `
          body { font-family: sans-serif; }
          table { border-collapse: collapse; margin-bottom: 24px; }
          th, td { border: 1px solid #ccc; padding: 4px 6px; font-variant-numeric: tabular-nums; }
          th { background: #f2f2f2; font-weight: bold; }
          .session-header td { border: none; padding: 2px 4px; }
          .best { background: #dff0d8; } /* light green */
          .worst { background: #f2dede; } /* light red */
          .avg { background: #d9edf7; } /* light blue */
          .lapPct { font-size: 10px; opacity: 0.7; white-space: nowrap; }
          .consistencyBar { font-family: monospace; font-size: 12px; }
          .analysis { font-size: 12px; }
        `;
        html += '</style></head><body>';
        sessions.forEach((session, sessionIndex) => {
          // Determine session date/time. Use first rider's first and last lap events for start/end.
          let startEpoch = Infinity;
          let endEpoch = -Infinity;
          (session.riders || []).forEach(r => {
            (r.events || []).filter(e => e.kind === 'LAP').forEach(ev => {
              if (typeof ev.stampEpoch === 'number') {
                if (ev.stampEpoch < startEpoch) startEpoch = ev.stampEpoch;
                if (ev.stampEpoch > endEpoch) endEpoch = ev.stampEpoch;
              }
            });
          });
          const sessionDate = session.dateISO ? new Date(session.dateISO) : (startEpoch !== Infinity ? new Date(startEpoch) : new Date());
          const fecha = formatDate(sessionDate);
          const horaInicio = startEpoch !== Infinity ? formatTime(new Date(startEpoch)) : '';
          const horaFin = endEpoch > startEpoch ? formatTime(new Date(endEpoch)) : '';
          // Header information rows
          html += '<table class="session-header">';
          html += `<tr><td colspan="2" style="font-size:20px;font-weight:700">ENDURO PRO</td></tr>`;
          html += `<tr><td>Fecha</td><td>${fecha}</td></tr>`;
          html += `<tr><td>Nombre de la tanda</td><td>${session.name || ('Tanda ' + (sessionIndex + 1))}</td></tr>`;
          html += `<tr><td>Hora de inicio</td><td>${horaInicio}</td></tr>`;
          html += `<tr><td>Hora de finalización</td><td>${horaFin}</td></tr>`;
          html += `<tr><td>Pista</td><td>${session.track || ''}</td></tr>`;
          html += `<tr><td>Seteo</td><td>${session.setupNotes || ''}</td></tr>`;
          html += `<tr><td>Cuenta inicial</td><td>${(session.prepSec ?? '') !== '' ? (session.prepSec + 's') : ''}</td></tr>`;
          html += `<tr><td>Gap</td><td>${(session.gapSec ?? '') !== '' ? (session.gapSec + 's') : ''}</td></tr>`;
          html += `<tr><td>Cantidad de parciales</td><td>${(session.sectors ?? '') !== '' ? String(session.sectors) : ''}</td></tr>`;
          html += `<tr><td>Notas</td><td>${session.notes || ''}</td></tr>`;
          html += '</table>';
          // Prepare rider lap data
          const ridersData = (session.riders || []).map(r => {
            const laps = (r.events || []).filter(e => e.kind === 'LAP');
            return { name: r.name || '', laps };
          });
          const maxLaps = Math.max(0, ...ridersData.map(r => r.laps.length));
          // Build table header
          html += '<table>';
          html += '<tr><th>Piloto</th>';
          for (let i = 1; i <= maxLaps; i++) {
            html += `<th>V${i}</th>`;
          }
          html += '<th>Promedio</th><th>Consistencia</th><th>Observación</th></tr>';
          // For each rider
          ridersData.forEach(r => {
            html += `<tr><td>${r.name}</td>`;
            const lapTimesMs = r.laps.map(l => l.lapNetMs);
            // Compute metrics
            let avg = null, best = null, worst = null;
            if (lapTimesMs.length) {
              best = Math.min(...lapTimesMs);
              worst = Math.max(...lapTimesMs);
              const sum = lapTimesMs.reduce((a, b) => a + b, 0);
              avg = sum / lapTimesMs.length;
            }
            // Insert lap cells up to maxLaps
            for (let i = 0; i < maxLaps; i++) {
              const lap = r.laps[i];
              if (!lap) {
                html += '<td></td>';
              } else {
                const timeStr = formatDuration(lap.lapNetMs);
                // Percentage difference relative to average
                let pctStr = '';
                let cellClass = '';
                if (avg) {
                  const ratio = (lap.lapNetMs - avg) / avg;
                  const pct = Math.abs(ratio) * 100;
                  const dir = ratio > 0 ? '▲' : ratio < 0 ? '▼' : '';
                  pctStr = `${dir} ${pct.toFixed(1)}%`;
                  // Determine class for best/worst/avg approximate
                  if (lap.lapNetMs === best) cellClass = 'best';
                  else if (lap.lapNetMs === worst) cellClass = 'worst';
                  else if (Math.abs(ratio) < 0.01) cellClass = 'avg';
                }
                html += `<td class="${cellClass}">${timeStr}<br><span class="lapPct">${pctStr}</span></td>`;
              }
            }
            // Average cell
            if (avg !== null) {
              html += `<td class="avg">${formatDuration(avg)}</td>`;
              // Consistency score: best/worst ratio scaled to 10
              const consistencyRatio = worst > 0 ? best / worst : 0;
              const consistencyScore = consistencyRatio * 10;
              // Build bar representation (10 characters)
              const filled = Math.round(consistencyScore);
              const empty = 10 - filled;
              let bar = '';
              bar = '█'.repeat(filled) + '░'.repeat(empty);
              html += `<td class="consistencyBar">${bar} ${consistencyScore.toFixed(1)}</td>`;
              // Observación / análisis
              let obs = '';
              if (lapTimesMs.length < 2) {
                obs = 'Datos insuficientes';
              } else {
                // Determine trend: first half vs second half
                const half = Math.floor(lapTimesMs.length / 2);
                const firstHalf = lapTimesMs.slice(0, half);
                const secondHalf = lapTimesMs.slice(half);
                const avgFirst = firstHalf.reduce((a, b) => a + b, 0) / firstHalf.length;
                const avgSecond = secondHalf.reduce((a, b) => a + b, 0) / secondHalf.length;
                const diffRatio = (avgSecond - avgFirst) / avgFirst;
                // Count slow laps relative to best
                const slowLaps = lapTimesMs.filter(t => (t - best) / best > 0.15).length;
                if (slowLaps >= 1) {
                  obs = 'Hubo una vuelta claramente fuera del promedio.';
                } else if (diffRatio > 0.05) {
                  obs = 'Caída de rendimiento al final.';
                } else if (diffRatio < -0.05) {
                  obs = 'Mejora hacia el final.';
                } else if (consistencyScore >= 9.0) {
                  obs = 'Muy constante.';
                } else if (consistencyScore >= 8.0) {
                  obs = 'Buen ritmo.';
                } else if (consistencyScore >= 6.0) {
                  obs = 'Ritmo irregular.';
                } else {
                  obs = 'Ritmo muy irregular.';
                }
              }
              html += `<td class="analysis">${obs}</td>`;
            } else {
              // No laps recorded
              html += '<td></td><td></td><td>Sin vueltas</td>';
            }
            html += '</tr>';
          });
          html += '</table>';
        });
        html += '</body></html>';
        // Create blob as Excel
        const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'jornada_enduro_pro_' + new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-') + '.xls';
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
        showToast('Informe PRO exportado');
      } btnStartPause.onclick = () => running ? pause() : start(); btnReset.onclick = resetTimeOnly; btnFinishSession.onclick = () => { if (confirm('¿Finalizar y guardar esta tanda?')) finishSession() }; btnClearAll.onclick = () => { if (confirm('¿Vaciar registros de la tanda actual?')) clearAll() }; btnClearDay.onclick = () => { if (confirm('¿Borrar la jornada actual completa?')) { dayHistory = []; saveDayHistory(); renderAll() } }; btnExport.onclick = exportExcelXlsx; btnToggleButtons.onclick = () => { document.body.classList.toggle('buttonsOnly'); save() }; btnAddRider.onclick = () => addRider(); btnResetNames.onclick = () => { if (confirm('¿Restaurar nombres por defecto?')) restoreNames() };[elPrepSec, elGap, elDblMs, elTrackName, elSessionName, elSessionNotes, elSetupNotes].forEach(el => el.oninput = () => { renderAll(); save() }); elSectors.onchange = () => { riders.forEach(r => { r.sectorIdx = 0; r.redo = [] }); renderAll(); save() }; loadDayHistory(); load(); ensureDefaults(); renderAll(); const ignitionScreen = document.getElementById('ignitionScreen');
      const ignitionStart = document.getElementById('ignitionStart');
      const splash = document.getElementById('splashScreen');
      let ignitionUsed = false;

      function finishSplashAfterStart(delayMs = 3600) {
        setTimeout(() => {
          if (splash) {
            splash.classList.add('hide');
            setTimeout(() => splash.classList.add('gone'), 450);
          }
        }, delayMs);
      }

      function startIgnitionExperience() {
        if (ignitionUsed) return;
        ignitionUsed = true;
        if (ignitionStart) ignitionStart.disabled = true;

        if (ignitionScreen) {
          ignitionScreen.classList.add('hide');
          setTimeout(() => ignitionScreen.classList.add('gone'), 380);
        }

        if (splash) {
          splash.classList.remove('hide', 'gone');
        }

        const braap = new Audio('audio/enduro_braap.mp3');
        braap.preload = 'auto';
        braap.volume = 0.9;

        let splashDelay = 3900;
        let finished = false;
        const finishOnce = () => {
          if (finished) return;
          finished = true;
          if (splash) {
            splash.classList.add('hide');
            setTimeout(() => splash.classList.add('gone'), 450);
          }
        };

        setTimeout(() => {
          braap.currentTime = 0;
          const playPromise = braap.play();
          if (playPromise && typeof playPromise.then === 'function') {
            playPromise
              .then(() => {
                braap.onended = finishOnce;
                setTimeout(finishOnce, 5200);
              })
              .catch(() => {
                finishSplashAfterStart(1800);
              });
          } else {
            setTimeout(finishOnce, splashDelay);
          }
        }, 220);
      }

      if (ignitionStart) {
        ignitionStart.addEventListener('click', startIgnitionExperience);
        ignitionStart.addEventListener('touchstart', (e) => {
          e.preventDefault();
          startIgnitionExperience();
        }, { passive: false });
      } else {
        finishSplashAfterStart(2000);
      }
      if ('serviceWorker' in navigator) {
        window.addEventListener('load', () => {
          navigator.serviceWorker.register('./sw.js').catch(() => {});
        });
      }
      window.addEventListener('visibilitychange', save);
      window.addEventListener('beforeunload', save);
      // Expose helpers for native XLSX export defined outside this module
      window.buildSessionsForExport = buildSessionsForExport; window.showToast = showToast;
    })();
