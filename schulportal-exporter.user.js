// ==UserScript==
// @name         Schulportal Hessen Universal Timetable Exporter (Print Fix)
// @namespace    https://github.com/0iy/schulportal-hessen-timetable-export
// @version      13.1
// @description  Generates valid ICS files, handles timezones, adds breaks, and fixes German character encoding & printing.
// @author       0iy
// @homepageURL  https://github.com/0iy/schulportal-hessen-timetable-export
// @supportURL   https://github.com/0iy/schulportal-hessen-timetable-export/issues
// @downloadURL  https://raw.githubusercontent.com/0iy/schulportal-hessen-timetable-export/main/schulportal-exporter.user.js
// @updateURL    https://raw.githubusercontent.com/0iy/schulportal-hessen-timetable-export/main/schulportal-exporter.user.js
// @match        https://start.schulportal.hessen.de/stundenplan.php*
// @grant        GM.xmlHttpRequest
// @connect      ferien-api.de
// @connect      feiertage-api.de
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    // --- Configuration & State ---
    const CONFIG = {
        POLL_INTERVAL_MS: 250,
        POLL_TIMEOUT_MS: 15000,
        EXPORT_TARGETS: [
            { id: 'own', printButtonSelector: '#printOwn', name: 'Personal' },
            { id: 'all', printButtonSelector: '#printAll', name: 'Overall' }
        ],
        PRINT_STYLE_ID: 'exporter-print-styles',
        STORAGE_KEY: 'timetable_class_selection'
    };

    const state = {
        isInitialized: false,
        currentExportTargetId: null,
        parsedTimetables: { own: [], all: [] }
    };

    const BUNDESLAENDER = {
        "HE": "Hessen", "BW": "Baden-Württemberg", "BY": "Bayern", "BE": "Berlin",
        "BB": "Brandenburg", "HB": "Bremen", "HH": "Hamburg", "MV": "Mecklenburg-Vorpommern",
        "NI": "Niedersachsen", "NW": "Nordrhein-Westfalen", "RP": "Rheinland-Pfalz", "SL": "Saarland",
        "SN": "Sachsen", "ST": "Sachsen-Anhalt", "SH": "Schleswig-Holstein", "TH": "Thüringen"
    };

    const log = {
        info: (msg) => console.log(`[Exporter] INFO: ${msg}`),
        error: (msg, err) => console.error(`[Exporter] ERROR: ${msg}`, err || '')
    };

    // --- LocalStorage ---
    function saveSelection() {
        const unchecked = Array.from(document.querySelectorAll('.exporter-class-checkbox:not(:checked)')).map(cb => cb.value);
        localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify(unchecked));
    }

    function loadSelection() {
        try {
            return new Set(JSON.parse(localStorage.getItem(CONFIG.STORAGE_KEY) || '[]'));
        } catch { return new Set(); }
    }

    // --- DOM & UI ---
    function injectModalCSS() {
        if (document.getElementById('exporter-modal-styles')) return;
        const css = `
            #exporter-modal-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.6); z-index: 10000; display: none; align-items: center; justify-content: center; font-family: sans-serif; }
            #exporter-modal-content { background: #fff; padding: 25px; border-radius: 8px; width: 90%; max-width: 800px; box-shadow: 0 10px 25px rgba(0,0,0,0.2); max-height: 90vh; display: flex; flex-direction: column; }
            #exporter-modal-header { font-size: 1.5em; font-weight: bold; border-bottom: 1px solid #eee; padding-bottom: 15px; margin-bottom: 20px; }
            #exporter-modal-body { overflow-y: auto; margin-bottom: 20px; flex: 1; }
            #exporter-class-selection { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 10px; }
            .class-item label { display: flex; align-items: center; cursor: pointer; padding: 5px; border-radius: 4px; transition: background 0.2s; }
            .class-item label:hover { background: #f0f0f0; }
            .class-item input { margin-right: 10px; transform: scale(1.2); }
            #exporter-ical-settings { margin-top: 20px; padding-top: 20px; border-top: 1px solid #eee; background: #fafafa; padding: 15px; border-radius: 5px; }
            #exporter-ical-settings h3 { margin-top: 0; font-size: 1.1em; color: #444; }
            .form-group { margin-bottom: 10px; display: flex; align-items: center; }
            .form-group label { width: 140px; font-weight: 600; }
            .form-group input[type="date"], .form-group select { padding: 6px; border: 1px solid #ccc; border-radius: 4px; flex: 1; }
            .checkbox-group { display: flex; align-items: center; cursor: pointer; }
            .checkbox-group input { margin-right: 10px; transform: scale(1.2); }
            #exporter-modal-footer { border-top: 1px solid #eee; padding-top: 15px; text-align: right; display: flex; justify-content: flex-end; gap: 10px; }
            .btn-exp { padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 600; text-decoration: none; color: #fff; display: inline-block; }
            .btn-cancel { background: #6c757d; }
            .btn-json { background: #007bff; }
            .btn-ical { background: #17a2b8; }
            .btn-print { background: #28a745; }
            .btn-exp:hover { opacity: 0.9; }
        `;
        const style = document.createElement('style');
        style.id = 'exporter-modal-styles';
        style.textContent = css;
        document.head.appendChild(style);
    }

    function createModal() {
        if (document.getElementById('exporter-modal-overlay')) return;
        const div = document.createElement('div');
        div.id = 'exporter-modal-overlay';
        div.innerHTML = `
            <div id="exporter-modal-content">
                <div id="exporter-modal-header">Export Timetable</div>
                <div id="exporter-modal-body">
                    <div id="exporter-class-selection"></div>
                    <div id="exporter-ical-settings">
                        <h3>Calendar Settings</h3>
                        <div class="form-group"><label>State:</label><select id="exporter-bundesland">${Object.entries(BUNDESLAENDER).map(([k,v]) => `<option value="${k}" ${k==='HE'?'selected':''}>${v}</option>`).join('')}</select></div>
                        <div class="form-group"><label>Start Date:</label><input type="date" id="exporter-start-date"></div>
                        <div class="form-group"><label>End Date:</label><input type="date" id="exporter-end-date"></div>
                        <div class="form-group">
                            <label>Extras:</label>
                            <label class="checkbox-group">
                                <input type="checkbox" id="exporter-include-breaks"> Include Breaks (Pauses)
                            </label>
                        </div>
                        <small style="color:#666;">Dates are auto-detected based on the current/next school year.</small>
                    </div>
                </div>
                <div id="exporter-modal-footer">
                    <button id="exp-cancel" class="btn-exp btn-cancel">Cancel</button>
                    <button id="exp-print" class="btn-exp btn-print">Print View</button>
                    <button id="exp-json" class="btn-exp btn-json">JSON</button>
                    <button id="exp-ical" class="btn-exp btn-ical">iCal (.ics)</button>
                </div>
            </div>`;
        document.body.appendChild(div);

        div.querySelector('#exp-cancel').onclick = hideModal;
        div.querySelector('#exp-json').onclick = handleJsonExport;
        div.querySelector('#exp-ical').onclick = handleICalExport;
        div.querySelector('#exp-print').onclick = handlePrint;
        div.querySelector('#exporter-bundesland').onchange = updateSchoolYearDates;
        div.onclick = (e) => { if (e.target.id === 'exporter-modal-overlay') hideModal(); };
    }

    function createExportButton(target) {
        const btn = document.createElement('a');
        btn.href = '#';
        btn.className = 'btn btn-primary exporter-btn';
        btn.style.marginLeft = '10px';
        btn.innerHTML = '<i class="fa fa-download"></i> Export';
        btn.onclick = (e) => {
            e.preventDefault();
            showModal(target.id, target.name);
        };
        return btn;
    }

    function checkAndInjectButtons() {
        CONFIG.EXPORT_TARGETS.forEach(t => {
            const printBtn = document.querySelector(t.printButtonSelector);
            if (printBtn && !printBtn.parentElement.querySelector('.exporter-btn')) {
                printBtn.parentElement.appendChild(createExportButton(t));
            }
        });
    }

    function showModal(targetId, targetName) {
        state.currentExportTargetId = targetId;
        document.getElementById('exporter-modal-header').textContent = `Export ${targetName} Timetable`;

        const uniquePairs = extractUniquePairs(targetId);
        const container = document.getElementById('exporter-class-selection');
        container.innerHTML = '';

        const unchecked = loadSelection();
        Array.from(uniquePairs.entries()).sort((a,b) => a[1].localeCompare(b[1])).forEach(([id, text]) => {
            const div = document.createElement('div');
            div.className = 'class-item';
            div.innerHTML = `<label><input type="checkbox" class="exporter-class-checkbox" value="${id}" ${!unchecked.has(id)?'checked':''}> ${text}</label>`;
            container.appendChild(div);
        });

        document.getElementById('exporter-modal-overlay').style.display = 'flex';
        updateSchoolYearDates();
    }

    function hideModal() {
        document.getElementById('exporter-modal-overlay').style.display = 'none';
    }

    // --- Parsing Logic ---
    function normalizeTable(table) {
        const rows = Array.from(table.querySelectorAll('tbody > tr'));
        const grid = [];
        rows.forEach((row, r) => {
            let c = 0;
            Array.from(row.cells).forEach(cell => {
                while (grid[r] && grid[r][c]) c++;
                const span = parseInt(cell.getAttribute('rowspan') || '1', 10);
                for (let i = 0; i < span; i++) {
                    if (!grid[r+i]) grid[r+i] = [];
                    grid[r+i][c] = cell;
                }
                c++;
            });
        });

        const newBody = document.createElement('tbody');
        grid.forEach(rowCells => {
            const tr = document.createElement('tr');
            rowCells.forEach(cell => {
                if(cell) tr.appendChild(cell.cloneNode(true));
                else tr.appendChild(document.createElement('td'));
            });
            newBody.appendChild(tr);
        });
        table.replaceChild(newBody, table.querySelector('tbody'));
    }

    function parseTimetables() {
        CONFIG.EXPORT_TARGETS.forEach(target => {
            const originalTable = document.querySelector(`#${target.id} .plan[data-date] table`);
            if (!originalTable) return;

            const table = originalTable.cloneNode(true);
            normalizeTable(table);

            const lessons = [];
            const headers = Array.from(table.querySelectorAll('thead th')).slice(1).map(th => th.textContent.trim());

            table.querySelectorAll('tbody tr').forEach(row => {
                const timeMeta = row.cells[0];
                const period = timeMeta.querySelector('b')?.textContent.trim();
                const time = timeMeta.querySelector('.VonBis small')?.textContent.trim();
                if (!period || !time) return;

                Array.from(row.cells).slice(1).forEach((cell, idx) => {
                    const day = headers[idx];
                    if (!day) return;

                    cell.querySelectorAll('.stunde').forEach(el => {
                        const subject = el.querySelector('b')?.textContent.trim();
                        const teacher = el.querySelector('small')?.textContent.trim();
                        // ROBUST ROOM EXTRACTION
                        let room = '';
                        el.childNodes.forEach(n => {
                            if(n.nodeType === 3 && n.textContent.trim()) room = n.textContent.trim();
                        });

                        if (subject) {
                            lessons.push({
                                day, period, time, subject, teacher, room,
                                uniqueId: `${subject}|${teacher}`,
                                isBreak: false
                            });
                        }
                    });
                });
            });
            state.parsedTimetables[target.id] = lessons;
        });
    }

    function extractUniquePairs(targetId) {
        const map = new Map();
        (state.parsedTimetables[targetId]||[]).forEach(l => {
            map.set(l.uniqueId, `${l.subject} (${l.teacher})`);
        });
        return map;
    }

    function getSelectedLessons(targetId) {
        const checkboxes = document.querySelectorAll('.exporter-class-checkbox:checked');
        const selectedIds = new Set(Array.from(checkboxes).map(c => c.value));
        return (state.parsedTimetables[targetId] || []).filter(l => selectedIds.has(l.uniqueId));
    }

    // --- Breaks / Pause Logic ---
    function injectBreaks(lessons) {
        const output = [...lessons];
        const dayMap = {};

        const toMins = (t) => {
            const [h, m] = t.split(':').map(Number);
            return h * 60 + m;
        };

        lessons.forEach(l => {
            if(!dayMap[l.day]) dayMap[l.day] = [];
            dayMap[l.day].push(l);
        });

        Object.keys(dayMap).forEach(day => {
            const sorted = dayMap[day].sort((a,b) => {
                return toMins(a.time.split(' - ')[0]) - toMins(b.time.split(' - ')[0]);
            });

            for(let i = 0; i < sorted.length - 1; i++) {
                const curr = sorted[i];
                const next = sorted[i+1];

                const currEnd = curr.time.split(' - ')[1]; // "08:45"
                const nextStart = next.time.split(' - ')[0]; // "08:50"

                const gapDiff = toMins(nextStart) - toMins(currEnd);

                if (gapDiff > 0) {
                    let name = "Pause";
                    if (gapDiff <= 10) name = "Wechselpause";
                    else if (gapDiff <= 25) name = "Große Pause";
                    else if (gapDiff >= 40) name = "Mittagspause";

                    output.push({
                        day: day,
                        time: `${currEnd} - ${nextStart}`,
                        subject: `${name} (${gapDiff} min)`,
                        teacher: '',
                        room: '',
                        uniqueId: `break-${day}-${currEnd}`,
                        isBreak: true
                    });
                }
            }
        });
        return output;
    }

    // --- Export Handlers ---
    function handleJsonExport() {
        saveSelection();
        const lessons = getSelectedLessons(state.currentExportTargetId);
        if(!lessons.length) return alert('No classes selected.');
        downloadFile(JSON.stringify(lessons, null, 2), 'timetable.json', 'application/json');
        hideModal();
    }

    function handlePrint() {
        saveSelection();
        const selectedIds = new Set(Array.from(document.querySelectorAll('.exporter-class-checkbox:checked')).map(cb => cb.value));
        if (selectedIds.size === 0) return alert('No classes selected.');

        hideModal();
        const targetId = state.currentExportTargetId;

        // FIXED PRINT CSS:
        let css = `
            @media print {
                @page { size: A4 landscape; margin: 0.5cm; }
                body { visibility: hidden; }
                #${targetId} {
                    visibility: visible;
                    position: absolute;
                    top: 0;
                    left: 0;
                    width: 100%;
                    margin: 0;
                    padding: 0;
                    background: white;
                }
                #${targetId} * { visibility: visible; }
                .print-buttons, .exporter-btn, .nav-tabs, h2+div, h2+div+small, .trenn { display: none !important; }
                .stunde { display: none !important; }
        `;

        // Reveal only selected classes
        selectedIds.forEach(id => {
            // Escape double quotes in ID if any
            const safeId = id.replace(/"/g, '\\"');
            css += `.stunde[data-id="${safeId}"] { display: block !important; }`;
        });

        css += `}`; // Close media query

        const style = document.createElement('style');
        style.id = CONFIG.PRINT_STYLE_ID;
        style.textContent = css;
        document.head.appendChild(style);

        window.print();
        setTimeout(() => { if(style) style.remove(); }, 1000);
    }

    async function handleICalExport() {
        saveSelection();
        let lessons = getSelectedLessons(state.currentExportTargetId);
        if(!lessons.length) return alert('No classes selected.');

        const includeBreaks = document.getElementById('exporter-include-breaks').checked;
        if (includeBreaks) {
            lessons = injectBreaks(lessons);
        }

        const startStr = document.getElementById('exporter-start-date').value;
        const endStr = document.getElementById('exporter-end-date').value;
        const land = document.getElementById('exporter-bundesland').value;

        if(!startStr || !endStr) return alert('Dates are required.');

        try {
            const startDate = new Date(startStr);
            const endDate = new Date(endStr);
            const years = new Set([startDate.getFullYear(), endDate.getFullYear()]);

            const holidays = await fetchHolidays(Array.from(years), land);
            const icsData = generateICS(lessons, startDate, endDate, holidays);

            downloadFile(icsData, 'timetable.ics', 'text/calendar');
            hideModal();
        } catch(e) {
            log.error('Export failed', e);
            alert('Export failed: ' + e.message);
        }
    }

    // --- API & Helpers ---
    async function fetchHolidays(years, land) {
        const dates = new Set();
        for (const year of years) {
            await new Promise((resolve) => {
                GM.xmlHttpRequest({
                    method: 'GET',
                    url: `https://feiertage-api.de/api/?jahr=${year}&nur_land=${land}`,
                    onload: (res) => {
                        if (res.status === 200) {
                            try {
                                const data = JSON.parse(res.responseText);
                                Object.values(data).forEach(h => dates.add(h.datum));
                            } catch(e) {}
                        }
                        resolve();
                    },
                    onerror: resolve
                });
            });
        }
        return dates;
    }

    async function updateSchoolYearDates() {
        const land = document.getElementById('exporter-bundesland').value;
        const sInput = document.getElementById('exporter-start-date');
        const eInput = document.getElementById('exporter-end-date');

        const now = new Date();
        let searchYear = now.getFullYear();
        if (now.getMonth() < 7) {
            searchYear--;
        }

        try {
            const hStart = await getSummerHoliday(land, searchYear);
            const hEnd = await getSummerHoliday(land, searchYear + 1);

            if(hStart && hEnd) {
                const start = new Date(hStart.end);
                start.setDate(start.getDate() + 1);

                const end = new Date(hEnd.start);
                end.setDate(end.getDate() - 1);

                sInput.value = start.toISOString().split('T')[0];
                eInput.value = end.toISOString().split('T')[0];
            }
        } catch(e) {
            log.error('Date fetch failed', e);
        }
    }

    function getSummerHoliday(land, year) {
        return new Promise((resolve, reject) => {
            GM.xmlHttpRequest({
                method: 'GET',
                url: `https://ferien-api.de/api/v1/holidays/${land}/${year}`,
                onload: (res) => {
                    if(res.status === 200) {
                        const data = JSON.parse(res.responseText);
                        const summer = data.find(h => h.name.toLowerCase().includes('sommerferien'));
                        resolve(summer);
                    } else resolve(null);
                },
                onerror: () => resolve(null)
            });
        });
    }

    // --- ICS Generator (RFC 5545 Compliant) ---
    function generateICS(lessons, startDate, endDate, holidays) {
        const mapDay = { 'Montag': 1, 'Dienstag': 2, 'Mittwoch': 3, 'Donnerstag': 4, 'Freitag': 5 };
        const formatTime = (isoDate, timeStr) => {
            const [h, m] = timeStr.split(':');
            const d = new Date(isoDate);
            d.setHours(parseInt(h,10), parseInt(m,10), 0);
            return d.getFullYear() +
                   String(d.getMonth()+1).padStart(2,'0') +
                   String(d.getDate()).padStart(2,'0') + 'T' +
                   String(d.getHours()).padStart(2,'0') +
                   String(d.getMinutes()).padStart(2,'0') + '00';
        };

        const ts = new Date().toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';
        const vTimezone = `
BEGIN:VTIMEZONE
TZID:Europe/Berlin
BEGIN:DAYLIGHT
TZOFFSETFROM:+0100
TZOFFSETTO:+0200
TZNAME:CEST
DTSTART:19700329T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:+0200
TZOFFSETTO:+0100
TZNAME:CET
DTSTART:19701025T030000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
END:STANDARD
END:VTIMEZONE`;

        let ics = [
            'BEGIN:VCALENDAR',
            'VERSION:2.0',
            'PRODID:-//SchulportalHessen//TimetableExporter//EN',
            'CALSCALE:GREGORIAN',
            vTimezone
        ];

        lessons.forEach((l, idx) => {
            const targetDay = mapDay[l.day];
            if (!targetDay) return;

            let firstDate = new Date(startDate);
            while(firstDate.getDay() !== targetDay) {
                firstDate.setDate(firstDate.getDate() + 1);
            }
            if(firstDate > endDate) return;

            const [startTime, endTime] = l.time.split(' - ');
            const startStr = formatTime(firstDate, startTime);
            const endStr = formatTime(firstDate, endTime);
            const untilStr = endDate.toISOString().replace(/[-:]/g, '').split('T')[0] + 'T235959Z';

            const exDates = [];
            holidays.forEach(hDateStr => {
                const hDate = new Date(hDateStr);
                if (hDate >= startDate && hDate <= endDate && hDate.getDay() === targetDay) {
                    exDates.push(`EXDATE;TZID=Europe/Berlin:${hDateStr.replace(/-/g,'')}T${startTime.replace(':','')}00`);
                }
            });

            // EMOJI & TEXT HANDLING
            const summary = l.isBreak ? `☕ ${l.subject}` : l.subject;
            const description = l.isBreak ? '' : `${l.subject} bei ${l.teacher}`;
            const location = l.isBreak ? '' : (l.room || '');

            ics.push(
                'BEGIN:VEVENT',
                `UID:${ts}-${idx}-${l.uniqueId.replace(/[^a-zA-Z0-9]/g,'')}@schulportal`,
                `DTSTAMP:${ts}`,
                `DTSTART;TZID=Europe/Berlin:${startStr}`,
                `DTEND;TZID=Europe/Berlin:${endStr}`,
                `SUMMARY:${summary}`,
                `LOCATION:${location}`,
                `DESCRIPTION:${description}`,
                `RRULE:FREQ=WEEKLY;UNTIL=${untilStr}`,
                ...exDates,
                'END:VEVENT'
            );
        });

        ics.push('END:VCALENDAR');
        return ics.join('\r\n');
    }

    function downloadFile(content, name, mime) {
        // FIX: Add Byte Order Mark (BOM) \uFEFF for explicit UTF-8 recognition by Excel/Calendar apps
        const blob = new Blob(["\uFEFF" + content], { type: mime + "; charset=utf-8" });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = name;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }

    // --- Init ---
    function init() {
        if (state.isInitialized) return;
        state.isInitialized = true;

        document.querySelectorAll('.stunde').forEach(el => {
            const s = el.querySelector('b')?.textContent.trim();
            const t = el.querySelector('small')?.textContent.trim();
            if(s && t) el.dataset.id = `${s}|${t}`;
        });

        injectModalCSS();
        createModal();
        parseTimetables();
        checkAndInjectButtons();

        new MutationObserver(checkAndInjectButtons).observe(document.body, {childList:true, subtree:true});
        log.info('Ready.');
    }

    const i = setInterval(() => {
        if(document.querySelector('.plan table')) {
            clearInterval(i);
            init();
        }
    }, CONFIG.POLL_INTERVAL_MS);
    setTimeout(() => clearInterval(i), CONFIG.POLL_TIMEOUT_MS);

})();