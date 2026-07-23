// ═══════════════════════════════════════════════════════════
//  BÁO CÁO KSAT - Google Sheets Backend (Apps Script API)
// ═══════════════════════════════════════════════════════════

// ★★★ THAY URL NÀY BẰNG URL APPS SCRIPT CỦA BẠN ★★★
const API_URL = 'https://script.google.com/macros/s/AKfycbz-tJf_LZ6p5vKe3bC2qY56XsUEUZrdWtw-YQsZkAIwl85Vl9wtv9XfPk5GVOaWSSqf/exec';

// ─── Google Sheets Store ───
class ReportStore {
    async save(report) {
        await fetch(API_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({ action: 'save', report })
        });
    }

    async getAll() {
        const res = await fetch(API_URL + '?action=list');
        return await res.json();
    }

    async get(id) {
        const res = await fetch(API_URL + '?action=get&id=' + encodeURIComponent(id));
        return await res.json();
    }

    async remove(id) {
        await fetch(API_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({ action: 'delete', id })
        });
    }

    async saveDriveItem(item) {
        try {
            const res = await fetch(API_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({ action: 'saveDriveItem', item })
            });
            return await res.json();
        } catch(e) { return null; }
    }

    async getDriveItems() {
        try {
            const res = await fetch(API_URL + '?action=listDriveItems');
            const data = await res.json();
            if (Array.isArray(data) && data.length > 0) return data;
        } catch(e) {}
        return null;
    }

    async removeDriveItem(id) {
        try {
            await fetch(API_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({ action: 'deleteDriveItem', id })
            });
        } catch(e) {}
    }
}

// ─── DOCX Parser (JSZip + XML) ───
async function parseDocx(file) {
    const zip = await JSZip.loadAsync(file);
    const xmlStr = await zip.file('word/document.xml').async('string');
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');
    const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

    const body = xmlDoc.getElementsByTagNameNS(ns, 'body')[0];
    const paragraphs = [];
    const tables = [];

    for (const child of body.children) {
        const tag = child.localName;
        if (tag === 'p') {
            const pData = parseParagraph(child, ns);
            if (pData.text.trim()) paragraphs.push(pData);
        } else if (tag === 'tbl') {
            const tData = parseTable(child, ns);
            if (tData.length > 0) tables.push(tData);
        }
    }

    return buildReport(paragraphs, tables, file.name);
}

function parseParagraph(pEl, ns) {
    let text = '';
    let totalRuns = 0;
    let boldRuns = 0;

    const runs = pEl.getElementsByTagNameNS(ns, 'r');
    for (const r of runs) {
        const tEls = r.getElementsByTagNameNS(ns, 't');
        let runText = '';
        for (const t of tEls) runText += t.textContent;
        text += runText;

        if (runText.trim()) {
            totalRuns++;
            const rPr = r.getElementsByTagNameNS(ns, 'rPr')[0];
            if (rPr) {
                const bEl = rPr.getElementsByTagNameNS(ns, 'b')[0];
                if (bEl) {
                    const val = bEl.getAttribute('w:val');
                    if (val !== '0' && val !== 'false') boldRuns++;
                }
            }
        }
    }

    return { text: text.trim(), bold: totalRuns > 0 && boldRuns >= totalRuns / 2 };
}

function parseTable(tblEl, ns) {
    const rows = tblEl.getElementsByTagNameNS(ns, 'tr');
    const result = [];
    for (const row of rows) {
        const cells = row.getElementsByTagNameNS(ns, 'tc');
        const rowData = [];
        for (const cell of cells) {
            let cellText = '';
            const ps = cell.getElementsByTagNameNS(ns, 'p');
            for (const p of ps) {
                const runs = p.getElementsByTagNameNS(ns, 'r');
                for (const r of runs) {
                    const ts = r.getElementsByTagNameNS(ns, 't');
                    for (const t of ts) cellText += t.textContent;
                }
                cellText += ' ';
            }
            rowData.push(cellText.trim());
        }
        result.push(rowData);
    }
    return result;
}

function isIgnoredNonTaskLine(text) {
    if (!text) return true;
    const clean = text.trim().toLowerCase();
    if (clean.length < 2) return true;

    const signoffPatterns = [
        /trân trọng/i,
        /kính gửi/i,
        /nơi nhận/i,
        /lưu:\s*/i,
        /lưu\s+vt/i,
        /^giám đốc/i,
        /^phó giám đốc/i,
        /^trưởng phòng/i,
        /^phó trưởng phòng/i,
        /^phụ trách/i,
        /^ký bởi/i,
        /^ký tên/i,
        /\.\/\./,
        /-\s*như trên/i,
        /-\s*lưu/i
    ];

    for (const pattern of signoffPatterns) {
        if (pattern.test(clean)) return true;
    }

    // Incident description / cause detail patterns (not planned tasks)
    if (clean.startsWith('do ') || clean.startsWith('do:')) return true; // e.g. "Do vi phạm..."
    if (clean.startsWith('khi ') && (clean.includes('thuộc pc') || clean.includes('trạm') || clean.includes('ngày') || clean.includes('công ty'))) return true; // e.g. "Khi thí nghiệm..."
    if (clean.includes('vi phạm khoảng cách') || clean.includes('tai nạn điện xảy ra')) return true;

    return false;
}

function buildReport(paragraphs, tables, filename) {
    const id = crypto.randomUUID();
    const report = {
        id, filename,
        upload_date: new Date().toISOString(),
        date_range: '',
        title: '',
        section1: [], section2: [], section3: [],
        table_data: []
    };

    let currentSection = 0;

    for (const p of paragraphs) {
        const lower = p.text.toLowerCase();

        if (currentSection === 0 && (lower.includes('(từ ngày') || lower.includes('từ ngày'))) {
            report.date_range = p.text;
        }

        if (/^i\.\s/i.test(lower) || lower.includes('báo cáo kết quả triển khai công tác')) {
            currentSection = 1;
            report.section1.push({ text: p.text, bold: true });
            continue;
        } else if (/^ii\.\s/i.test(lower) || lower.includes('kế hoạch công tác từ ngày')) {
            currentSection = 2;
            report.section2.push({ text: p.text, bold: true });
            continue;
        } else if (/^iii\.\s/i.test(lower) || lower.includes('kiến nghị về công tác')) {
            currentSection = 3;
            report.section3.push({ text: p.text, bold: true });
            continue;
        }

        if (isIgnoredNonTaskLine(p.text)) continue;

        if (currentSection === 1) report.section1.push(p);
        else if (currentSection === 2) report.section2.push(p);
        else if (currentSection === 3) report.section3.push(p);
    }

    for (const table of tables) {
        if (table.length > 1 && table[0].length >= 7) {
            const h0 = table[0][0].toLowerCase();
            const h1 = table[0][1].toLowerCase();
            if (h0.includes('stt') && h1.includes('đơn vị')) {
                for (let i = 1; i < table.length; i++) {
                    const r = table[i];
                    if (r.length >= 7 && /^\d+$/.test(r[0].trim())) {
                        const toInt = s => { const n = parseInt(s.replace(/\D/g, '')); return isNaN(n) ? 0 : n; };
                        report.table_data.push({
                            stt: r[0].trim(),
                            don_vi: r[1].trim(),
                            ke_hoach: toInt(r[2]),
                            trung_binh: toInt(r[3]),
                            thuc_hien: toInt(r[4]),
                            luy_ke: toInt(r[5]),
                            con_lai: toInt(r[6])
                        });
                    }
                }
                break;
            }
        }
    }

    report.title = report.date_range || `Báo cáo ngày ${new Date().toLocaleDateString('vi-VN')}`;
    return report;
}

// ═══════════════════════════════════════════════════════════
//  UI CONTROLLER
// ═══════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', async () => {
    const $ = id => document.getElementById(id);
    const fileInput = $('fileInput');
    const uploadForm = $('uploadForm');
    const uploadBtn = $('uploadBtn');
    const uploadText = $('uploadText');
    const uploadStatus = $('uploadStatus');
    const reportList = $('reportList');
    const reportCount = $('reportCount');
    const compareBtn = $('compareBtn');
    const closeCompare = $('closeCompare');
    const reportView = $('reportView');
    const compareArea = $('compareArea');
    const emptyState = $('emptyState');
    const mainTitle = $('mainTitle');
    const subTitle = $('subTitle');
    const deleteBtn = $('deleteBtn');
    const kpiCards = $('kpiCards');
    const tableSection = $('tableSection');
    const dropZone = $('dropZone');

    let reports = [];
    let currentReportId = null;
    let selectedForCompare = new Set();
    let charts = {};
    let activeComparisonData = null;

    const store = new ReportStore();

    // ─── THEME TOGGLE ───
    const themeToggle = $('themeToggle');
    const themeToggleMobile = $('themeToggleMobile');
    const savedTheme = localStorage.getItem('atvs-theme') || 'dark';
    if (savedTheme === 'light') document.documentElement.setAttribute('data-theme', 'light');

    function toggleTheme() {
        const isLight = document.documentElement.getAttribute('data-theme') === 'light';
        if (isLight) {
            document.documentElement.removeAttribute('data-theme');
            localStorage.setItem('atvs-theme', 'dark');
        } else {
            document.documentElement.setAttribute('data-theme', 'light');
            localStorage.setItem('atvs-theme', 'light');
        }
    }
    themeToggle.addEventListener('click', toggleTheme);
    if (themeToggleMobile) themeToggleMobile.addEventListener('click', toggleTheme);

    // ─── MOBILE SIDEBAR ───
    const sidebar = $('sidebar');
    const menuToggle = $('menuToggle');
    const sidebarOverlay = $('sidebarOverlay');

    function openSidebar() {
        sidebar.classList.add('open');
        sidebarOverlay.classList.add('active');
        document.body.style.overflow = 'hidden';
    }
    function closeSidebar() {
        sidebar.classList.remove('open');
        sidebarOverlay.classList.remove('active');
        document.body.style.overflow = '';
    }
    if (menuToggle) menuToggle.addEventListener('click', openSidebar);
    if (sidebarOverlay) sidebarOverlay.addEventListener('click', closeSidebar);

    // ─── INIT ───
    await fetchReports();

    // ─── DRAG & DROP ───
    ['dragenter','dragover'].forEach(e => dropZone.addEventListener(e, ev => { ev.preventDefault(); dropZone.classList.add('dragover'); }));
    ['dragleave','drop'].forEach(e => dropZone.addEventListener(e, ev => { ev.preventDefault(); dropZone.classList.remove('dragover'); }));
    dropZone.addEventListener('drop', ev => {
        const file = ev.dataTransfer.files[0];
        if (file && file.name.endsWith('.docx')) {
            fileInput.files = ev.dataTransfer.files;
            uploadText.textContent = file.name;
            uploadBtn.disabled = false;
        }
    });

    fileInput.addEventListener('change', e => {
        if (e.target.files.length) {
            uploadText.textContent = e.target.files[0].name;
            uploadBtn.disabled = false;
        } else {
            uploadText.textContent = 'Kéo thả hoặc chọn file .docx';
            uploadBtn.disabled = true;
        }
    });

    uploadForm.addEventListener('submit', async e => {
        e.preventDefault();
        const file = fileInput.files[0];
        if (!file) return;

        uploadBtn.disabled = true;
        uploadBtn.innerHTML = '<span class="spinner"></span> Đang xử lý...';
        uploadStatus.textContent = '';
        uploadStatus.className = 'status-msg';

        try {
            const report = await parseDocx(file);

            uploadStatus.textContent = '⏳ Đang lưu lên server...';
            uploadStatus.classList.add('status-success');
            await store.save(report);

            uploadStatus.textContent = '✓ Tải lên thành công!';
            fileInput.value = '';
            uploadText.textContent = 'Kéo thả hoặc chọn file .docx';
            await fetchReports();
            loadReport(report.id);
        } catch (err) {
            console.error(err);
            uploadStatus.textContent = '✗ Lỗi: ' + err.message;
            uploadStatus.classList.add('status-error');
        } finally {
            uploadBtn.disabled = false;
            uploadBtn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16"><path d="M5 12h14M12 5l7 7-7 7"/></svg> Tải Lên';
            setTimeout(() => { uploadStatus.textContent = ''; }, 4000);
        }
    });

    // ─── TABS ───
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            btn.classList.add('active');
            $(btn.dataset.tab).classList.add('active');
        });
    });

    // ─── DELETE ───
    deleteBtn.addEventListener('click', async () => {
        if (!currentReportId || !confirm('Bạn có chắc muốn xóa báo cáo này?')) return;
        uploadStatus.textContent = '⏳ Đang xóa...';
        await store.remove(currentReportId);
        currentReportId = null;
        showView('empty');
        await fetchReports();
        uploadStatus.textContent = '';
    });

    // ─── COMPARE ───
    compareBtn.addEventListener('click', () => {
        if (selectedForCompare.size >= 2) loadComparison(Array.from(selectedForCompare));
    });
    closeCompare.addEventListener('click', () => {
        if (currentReportId) showView('report');
        else showView('empty');
    });

    // ─── DATA ───
    async function fetchReports() {
        try {
            reports = await store.getAll();
            reportCount.textContent = reports.length;
            renderReportList();
        } catch (err) {
            console.error(err);
            reportList.innerHTML = '<div class="empty-timeline">Lỗi kết nối server</div>';
        }
    }

    function isAuthorizedAccess() {
        const urlParams = new URLSearchParams(window.location.search);
        const sourceParam = urlParams.get('source') || urlParams.get('from') || urlParams.get('mode') || urlParams.get('ref') || urlParams.get('access');
        if (sourceParam && (sourceParam.toLowerCase().includes('draftpcvt') || sourceParam.toLowerCase().includes('ksat') || sourceParam === 'admin')) {
            sessionStorage.setItem('ksat_authorized_compare', 'true');
            return true;
        }
        const referrer = document.referrer || '';
        if (referrer.toLowerCase().includes('draftpcvt.vercel.app') || referrer.toLowerCase().includes('draftpcvt')) {
            sessionStorage.setItem('ksat_authorized_compare', 'true');
            return true;
        }
        if (sessionStorage.getItem('ksat_authorized_compare') === 'true') {
            return true;
        }
        return false;
    }

    function renderReportList() {
        const canCompare = isAuthorizedAccess();

        const compareActionsContainer = $('compareActionsContainer');
        if (compareActionsContainer) {
            if (canCompare) compareActionsContainer.classList.remove('hidden');
            else compareActionsContainer.classList.add('hidden');
        }

        const btnOpenReminder = $('btnOpenReminder');
        if (btnOpenReminder) {
            if (canCompare) btnOpenReminder.classList.remove('hidden');
            else btnOpenReminder.classList.add('hidden');
        }

        if (!reports.length) {
            reportList.innerHTML = '<div class="empty-timeline">Chưa có báo cáo</div>';
            return;
        }
        reportList.innerHTML = '';
        reports.forEach(r => {
            const div = document.createElement('div');
            div.className = `report-item ${r.id === currentReportId ? 'active' : ''}`;

            div.innerHTML = `<span class="dot"></span><div class="report-info">
                <div class="report-title">${escapeHtml(r.title || 'Không tên')}</div>
                <div class="report-date">${new Date(r.upload_date).toLocaleString('vi-VN')}</div>
            </div>`;

            if (canCompare) {
                const cb = document.createElement('input');
                cb.type = 'checkbox'; cb.className = 'report-checkbox';
                cb.checked = selectedForCompare.has(r.id);
                cb.onclick = e => {
                    e.stopPropagation();
                    cb.checked ? selectedForCompare.add(r.id) : selectedForCompare.delete(r.id);
                    compareBtn.disabled = selectedForCompare.size < 2;
                    compareBtn.querySelector('svg').nextSibling.textContent = ` So sánh (${selectedForCompare.size})`;
                };
                div.prepend(cb);
            }

            div.onclick = () => { closeSidebar(); loadReport(r.id); };
            reportList.appendChild(div);
        });
    }

    async function loadReport(id) {
        try {
            mainTitle.textContent = '⏳ Đang tải...';
            subTitle.textContent = '';
            showView('report');

            const data = await store.get(id);
            if (data.error) throw new Error(data.error);
            currentReportId = id;
            renderReportList();

            mainTitle.textContent = data.title;
            subTitle.textContent = `Ngày tải lên: ${new Date(data.upload_date).toLocaleString('vi-VN')}`;

            renderKPI(data.table_data);
            renderStructuredSection('section1Content', data.section1, 'section1');
            renderStructuredSection('section2Content', data.section2, 'section2');
            renderStructuredSection('section3Content', data.section3, 'section3');
            renderTable(data.table_data);
        } catch (e) { console.error(e); alert('Không thể tải báo cáo: ' + e.message); }
    }

    // ─── KPI CARDS ───
    function renderKPI(tableData) {
        if (!tableData || !tableData.length) { kpiCards.innerHTML = ''; return; }
        const totalKH = tableData.reduce((s, r) => s + r.ke_hoach, 0);
        const totalTH = tableData.reduce((s, r) => s + r.thuc_hien, 0);
        const totalLK = tableData.reduce((s, r) => s + r.luy_ke, 0);
        const totalCL = tableData.reduce((s, r) => s + r.con_lai, 0);
        const pct = totalKH > 0 ? ((totalLK / totalKH) * 100).toFixed(1) : 0;

        kpiCards.innerHTML = `
            <div class="kpi-section-title">📊 Tổng hợp Phiếu Khảo sát, tư vấn điện gia đình</div>
            <div class="kpi-card kpi-accent"><div class="kpi-label">🎯 Kế hoạch cả năm</div><div class="kpi-value">${totalKH.toLocaleString()}</div><div class="kpi-sub">Tổng chỉ tiêu phiếu KSTV được giao</div></div>
            <div class="kpi-card kpi-green"><div class="kpi-label">📝 Thực hiện trong tuần</div><div class="kpi-value">${totalTH.toLocaleString()}</div><div class="kpi-sub">Số phiếu KSTV hoàn thành tuần này</div></div>
            <div class="kpi-card kpi-purple"><div class="kpi-label">📈 Lũy kế đến hiện tại</div><div class="kpi-value">${totalLK.toLocaleString()}</div><div class="kpi-sub">Đạt ${pct}% so với kế hoạch năm</div></div>
            <div class="kpi-card kpi-amber"><div class="kpi-label">⏳ Còn lại phải hoàn thành</div><div class="kpi-value">${totalCL.toLocaleString()}</div><div class="kpi-sub">Số phiếu KSTV cần làm thêm</div></div>
        `;
    }

    // ─── STRUCTURED CONTENT ───
    const sectionMeta = {
        section1: {
            groupKeywords: [
                { match: 'an toàn lao động', icon: '🛡️', theme: 'green', label: 'An toàn lao động' },
                { match: 'bảo vệ an toàn công trình', icon: '🏗️', theme: 'accent', label: 'Bảo vệ an toàn công trình điện lực' },
                { match: 'pccc', icon: '🧯', theme: 'red', label: 'Phòng cháy chữa cháy & CNCH' },
            ],
            fallback: { icon: '📋', theme: 'accent', label: 'Nội dung' }
        },
        section2: {
            groupKeywords: [{ match: 'kế hoạch', icon: '📅', theme: 'cyan', label: 'Kế hoạch công tác' }],
            fallback: { icon: '📅', theme: 'cyan', label: 'Kế hoạch' }
        },
        section3: {
            groupKeywords: [
                { match: 'dụng cụ an toàn', icon: '🔧', theme: 'amber', label: 'Dụng cụ an toàn' },
                { match: 'trang bị thêm dây an toàn', icon: '🧤', theme: 'green', label: 'Trang bị dây an toàn, sào, găng' },
                { match: 'kìm ép', icon: '⚙️', theme: 'purple', label: 'Kìm ép, dụng cụ cắt cáp' },
                { match: 'app công trường', icon: '📱', theme: 'accent', label: 'App Công trường' },
                { match: 'cskh', icon: '🎧', theme: 'cyan', label: 'TT CSKH & ITSM' },
                { match: 'chu trình điều phối', icon: '🔄', theme: 'red', label: 'Chu trình điều phối PTT, PCT' },
            ],
            fallback: { icon: '⚠️', theme: 'amber', label: 'Kiến nghị' }
        }
    };

    function renderStructuredSection(elementId, dataArray, sectionKey) {
        const el = $(elementId);
        const filteredArray = (dataArray || []).filter(item => item && item.text && !isIgnoredNonTaskLine(item.text));
        if (!filteredArray || !filteredArray.length) {
            el.innerHTML = '<div class="content-group"><div class="group-body open"><div class="content-item">Không có nội dung.</div></div></div>';
            return;
        }
        const meta = sectionMeta[sectionKey];
        const groups = groupItems(filteredArray, meta);

        el.innerHTML = groups.map((g, gi) => {
            const items = g.items.map(item => {
                const isResponse = item.text.startsWith('→') || item.text.startsWith('->');
                return `<div class="content-item ${isResponse ? 'is-response' : ''}">
                    <span class="bullet"></span><span>${escapeHtml(item.text)}</span>
                </div>`;
            }).join('');

            return `<div class="content-group theme-${g.theme}">
                <div class="group-header ${gi === 0 ? 'expanded' : ''}" onclick="toggleGroup(this)">
                    <div class="group-icon">${g.icon}</div>
                    <div class="group-title">${escapeHtml(g.label)}</div>
                    <span class="group-count">${g.items.length} mục</span>
                    <svg class="group-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 9l6 6 6-6"/></svg>
                </div>
                <div class="group-body ${gi === 0 ? 'open' : ''}">${items}</div>
            </div>`;
        }).join('');
    }

    function groupItems(items, meta) {
        const groups = [];
        let current = null;
        items.forEach(item => {
            if (item.bold) {
                const lower = item.text.toLowerCase();
                let matched = meta.groupKeywords.find(k => lower.includes(k.match));
                if (matched) {
                    current = { icon: matched.icon, theme: matched.theme, label: item.text, items: [] };
                    groups.push(current);
                    return;
                }
                if (current) { current.items.push(item); }
                else {
                    if (lower.includes('báo cáo kết quả') || lower.includes('kế hoạch công tác') || lower.includes('kiến nghị về công tác')) return;
                    current = { ...meta.fallback, label: item.text, items: [] };
                    groups.push(current);
                }
            } else {
                if (!current) { current = { ...meta.fallback, items: [] }; groups.push(current); }
                current.items.push(item);
            }
        });
        return groups.filter(g => g.items.length > 0);
    }

    window.toggleGroup = function(header) {
        header.classList.toggle('expanded');
        header.nextElementSibling.classList.toggle('open');
    };

    // ─── TABLE ───
    function renderTable(tableData) {
        const tbody = document.querySelector('#dataTable tbody');
        if (!tableData || !tableData.length) { tableSection.classList.add('hidden'); return; }
        tableSection.classList.remove('hidden');
        tbody.innerHTML = tableData.map(row => {
            const pct = row.ke_hoach > 0 ? ((row.luy_ke / row.ke_hoach) * 100) : 0;
            const color = pct > 30 ? 'var(--green)' : pct > 15 ? 'var(--amber)' : 'var(--red)';
            return `<tr>
                <td class="num">${row.stt}</td>
                <td>${escapeHtml(row.don_vi)}</td>
                <td class="num">${row.ke_hoach.toLocaleString()}</td>
                <td class="num">${row.trung_binh.toLocaleString()}</td>
                <td class="num" style="color:var(--accent);font-weight:600">${row.thuc_hien.toLocaleString()}</td>
                <td class="num" style="font-weight:600">${row.luy_ke.toLocaleString()}</td>
                <td class="num" style="color:var(--amber)">${row.con_lai.toLocaleString()}</td>
                <td class="progress-cell">
                    <div class="progress-bar-bg"><div class="progress-bar-fill" style="width:${pct.toFixed(1)}%;background:${color}"></div></div>
                    <div class="progress-pct" style="color:${color}">${pct.toFixed(1)}%</div>
                </td>
            </tr>`;
        }).join('');
    }

    // ─── COMPARE ───
    async function loadComparison(ids) {
        if (!isAuthorizedAccess()) {
            alert('Tính năng so sánh chỉ khả dụng khi truy cập từ trang Quản lý (DraftPCVT)!');
            showView('empty');
            return;
        }
        try {
            const allData = [];
            for (const id of ids) { const d = await store.get(id); if (d && !d.error) allData.push(d); }
            allData.sort((a, b) => new Date(a.upload_date) - new Date(b.upload_date));
            showView('compare');
            $('compareSubtitle').textContent = `${allData.length} tuần được chọn`;

            const labels = allData.map(d => { const t = d.title || ''; return t.length > 35 ? t.substring(0, 35) + '…' : t; });
            const weeklyData = allData.map(r => (r.table_data || []).reduce((s, row) => s + row.thuc_hien, 0));
            const cumulativeData = allData.map(r => (r.table_data || []).reduce((s, row) => s + row.luy_ke, 0));

            // Summary paragraph (above charts)
            if (allData.length >= 2) {
                const prevWeek = allData[allData.length - 2];
                const currWeek = allData[allData.length - 1];
                activeComparisonData = { prevWeek, currWeek };
                renderCompareSummary(prevWeek, currWeek);
            } else {
                activeComparisonData = null;
            }

            renderChart('weeklyChart', 'Tổng TH trong tuần', labels, weeklyData, '#3b82f6');
            renderChart('cumulativeChart', 'Tổng Lũy kế', labels, cumulativeData, '#10b981');
            renderDeltas(allData);

            // Plan vs Result + Detail table
            if (allData.length >= 2) {
                const prevWeek = allData[allData.length - 2];
                const currWeek = allData[allData.length - 1];
                renderPlanVsResult(prevWeek, currWeek);
                renderPlanAnalysis(prevWeek, currWeek);
                renderDetailComparison(prevWeek, currWeek);
            }
        } catch (e) { console.error(e); alert('Lỗi khi so sánh'); }
    }

    function renderPlanAnalysis(prevWeek, currWeek) {
        const container = $('planAnalysisContent');
        const planItems = (prevWeek.section2 || []);
        const resultItems = (currWeek.section1 || []);

        // Extract text arrays (non-heading items only for matching)
        const resultTexts = resultItems.filter(r => !r.bold).map(r => r.text.toLowerCase());
        const allResultText = resultTexts.join(' ');

        // Extract keywords from a text string
        function extractKeywords(text) {
            const stopWords = new Set(['và','các','của','cho','với','trong','từ','về','theo','đã','được','có','để','tại','do','là','một','này','đó','trên','khi','sẽ','đến','ra','lên','không','những','công','tác','việc','tuần','ngày','năm','tháng']);
            return text.toLowerCase()
                .replace(/[.,;:!?()"']/g, ' ')
                .split(/\s+/)
                .filter(w => w.length > 2 && !stopWords.has(w) && !/^\d+$/.test(w));
        }

        // Calculate similarity score between plan item and all results
        function findBestMatch(planText) {
            const planKw = extractKeywords(planText);
            if (planKw.length === 0) return { score: 0, matches: [] };

            let matchedKeywords = 0;
            const matchingResults = [];

            // Check each keyword against all result text
            for (const kw of planKw) {
                if (allResultText.includes(kw)) matchedKeywords++;
            }

            // Find specific matching result lines
            for (const rt of resultTexts) {
                let lineMatch = 0;
                for (const kw of planKw) {
                    if (rt.includes(kw)) lineMatch++;
                }
                if (lineMatch >= Math.min(2, planKw.length * 0.3)) {
                    matchingResults.push(resultItems.find(r => r.text.toLowerCase() === rt));
                }
            }

            const score = planKw.length > 0 ? matchedKeywords / planKw.length : 0;
            return { score, matches: matchingResults.slice(0, 3) };
        }

        // Analyze each plan item
        const analysis = [];
        let currentCategory = '';

        for (const item of planItems) {
            if (item.bold) {
                currentCategory = item.text;
                analysis.push({ type: 'category', text: item.text });
                continue;
            }

            // Skip very short, empty, signoff lines, or sub-bullets starting with "+"
            if (!item.text || item.text.trim().length < 5 || isIgnoredNonTaskLine(item.text) || item.text.trim().startsWith('+')) continue;

            const { score, matches } = findBestMatch(item.text);
            let status, statusLabel, statusBadge, statusClass;

            if (score >= 0.5) {
                status = 'done'; statusLabel = '✅ Đã triển khai'; statusBadge = 'pa-status-done'; statusClass = 'pa-item-done';
            } else if (score >= 0.25) {
                status = 'partial'; statusLabel = '⚠️ Triển khai một phần'; statusBadge = 'pa-status-partial'; statusClass = 'pa-item-partial';
            } else {
                status = 'none'; statusLabel = '❌ Chưa triển khai'; statusBadge = 'pa-status-none'; statusClass = 'pa-item-none';
            }

            analysis.push({
                type: 'item', text: item.text, status, statusLabel, statusBadge, statusClass, matches, score
            });
        }

        // Count stats
        const items = analysis.filter(a => a.type === 'item');
        const doneCount = items.filter(a => a.status === 'done').length;
        const partialCount = items.filter(a => a.status === 'partial').length;
        const noneCount = items.filter(a => a.status === 'none').length;
        const totalItems = items.length;
        const completionRate = totalItems > 0 ? ((doneCount / totalItems) * 100).toFixed(0) : 0;

        // Build HTML
        let html = `
            <div class="plan-analysis-stats">
                <div class="pa-stat pa-done"><span class="pa-count">${doneCount}</span> Đã triển khai</div>
                <div class="pa-stat pa-partial"><span class="pa-count">${partialCount}</span> Một phần</div>
                <div class="pa-stat pa-none"><span class="pa-count">${noneCount}</span> Chưa triển khai</div>
                <div class="pa-stat" style="background:var(--accent-soft);color:var(--accent)">Hoàn thành: <span class="pa-count">${completionRate}%</span></div>
            </div>
            <div class="plan-analysis-list">
        `;

        for (const a of analysis) {
            if (a.type === 'category') {
                html += `<div class="pa-category">${escapeHtml(a.text)}</div>`;
            } else {
                html += `<div class="pa-item ${a.statusClass}">
                    <div class="pa-item-header">
                        <span class="pa-status-badge ${a.statusBadge}">${a.statusLabel}</span>
                    </div>
                    <div class="pa-plan-text">📋 ${escapeHtml(a.text)}</div>
                    ${a.matches.length > 0 ? `<div class="pa-match-text">
                        <strong>Kết quả tương ứng:</strong><br>
                        ${a.matches.map(m => '• ' + escapeHtml(m.text)).join('<br>')}
                    </div>` : `<div class="pa-match-text" style="color:var(--red);border-left-color:var(--red)">
                        Không tìm thấy kết quả tương ứng trong báo cáo tuần này.
                    </div>`}
                </div>`;
            }
        }

        html += '</div>';
        container.innerHTML = html;
    }

    function renderCompareSummary(prev, curr) {
        const el = $('compareSummary');
        const prevTitle = prev.title || 'Tuần trước';
        const currTitle = curr.title || 'Tuần này';

        const sumVal = (data, key) => (data.table_data || []).reduce((s, r) => s + r[key], 0);

        const prevTH = sumVal(prev, 'thuc_hien');
        const currTH = sumVal(curr, 'thuc_hien');
        const prevLK = sumVal(prev, 'luy_ke');
        const currLK = sumVal(curr, 'luy_ke');
        const prevCL = sumVal(prev, 'con_lai');
        const currCL = sumVal(curr, 'con_lai');
        const prevKH = sumVal(prev, 'ke_hoach');
        const currKH = sumVal(curr, 'ke_hoach');

        const diffTH = currTH - prevTH;
        const diffLK = currLK - prevLK;
        const pctChangeTH = prevTH !== 0 ? ((diffTH / prevTH) * 100).toFixed(1) : '—';
        const pctProgress = currKH > 0 ? ((currLK / currKH) * 100).toFixed(1) : '—';

        // Determine trend
        const trendTH = diffTH > 0 ? 'tăng' : diffTH < 0 ? 'giảm' : 'không đổi';
        const trendClsTH = diffTH > 0 ? 'hl-green' : diffTH < 0 ? 'hl-red' : 'hl-amber';

        // Find best and worst performing units
        const currData = curr.table_data || [];
        let bestUnit = '', worstUnit = '', bestVal = -Infinity, worstVal = Infinity;
        currData.forEach(r => {
            if (r.thuc_hien > bestVal) { bestVal = r.thuc_hien; bestUnit = r.don_vi; }
            if (r.thuc_hien < worstVal) { worstVal = r.thuc_hien; worstUnit = r.don_vi; }
        });

        el.innerHTML = `
            <div class="summary-title">📝 Tổng quan so sánh</div>
            <p>
                So sánh giữa <span class="hl-accent">${escapeHtml(prevTitle)}</span> và <span class="hl-accent">${escapeHtml(currTitle)}</span>:
            </p>
            <p>
                Tuần này, tổng số phiếu KSTV thực hiện là <span class="highlight">${currTH.toLocaleString()} phiếu</span>,
                <span class="${trendClsTH}">${trendTH} ${Math.abs(diffTH).toLocaleString()} phiếu</span>
                (${diffTH >= 0 ? '+' : ''}${pctChangeTH}%) so với tuần trước (${prevTH.toLocaleString()} phiếu).
            </p>
            <p>
                Lũy kế đến hiện tại đạt <span class="hl-green">${currLK.toLocaleString()} phiếu</span>,
                tương đương <span class="highlight">${pctProgress}%</span> kế hoạch cả năm (${currKH.toLocaleString()} phiếu).
                Còn lại phải hoàn thành: <span class="hl-amber">${currCL.toLocaleString()} phiếu</span>.
            </p>
            ${bestUnit ? `<p>
                Đơn vị thực hiện tốt nhất tuần này: <span class="hl-green">${escapeHtml(bestUnit)}</span> (${bestVal.toLocaleString()} phiếu).
                ${worstUnit && worstUnit !== bestUnit ? `Đơn vị cần cải thiện: <span class="hl-red">${escapeHtml(worstUnit)}</span> (${worstVal.toLocaleString()} phiếu).` : ''}
            </p>` : ''}
            <div style="margin-top: 1rem; padding-top: 0.75rem; border-top: 1px dashed var(--border);">
                <button id="btnQuickReminder" class="btn primary-btn btn-reminder">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16"><path d="M22 17H2a3 3 0 0 0 3-3V9a7 7 0 0 1 14 0v5a3 3 0 0 0 3 3zm-8.27 4a2 2 0 0 1-3.46 0"/></svg>
                    <span>📢 Xuất thông báo nhắc nhở KSAT</span>
                </button>
            </div>
        `;

        const btnQuick = $('btnQuickReminder');
        if (btnQuick) btnQuick.addEventListener('click', openReminderModal);
    }

    function renderPlanVsResult(prevWeek, currWeek) {
        const container = $('planVsResultContent');
        const prevTitle = prevWeek.title || 'Tuần trước';
        const currTitle = currWeek.title || 'Tuần này';

        function renderItems(items) {
            const valid = (items || []).filter(item => item.bold || (!isIgnoredNonTaskLine(item.text) && item.text.trim().length >= 5));
            if (!valid.length) return '<div class="compare-item" style="color:var(--text-tertiary)">Không có dữ liệu</div>';
            return valid.map(item => {
                if (item.bold) {
                    return `<div class="compare-item ci-heading"><span>${escapeHtml(item.text)}</span></div>`;
                }
                return `<div class="compare-item"><span class="ci-bullet"></span><span>${escapeHtml(item.text)}</span></div>`;
            }).join('');
        }

        container.innerHTML = `
            <div class="compare-column">
                <div class="compare-column-header">
                    <span class="col-icon">📋</span>
                    <span>Kế hoạch tuần trước</span>
                    <span class="col-week">${escapeHtml(prevTitle)}</span>
                </div>
                ${renderItems(prevWeek.section2)}
            </div>
            <div class="compare-column">
                <div class="compare-column-header">
                    <span class="col-icon">✅</span>
                    <span>Kết quả thực hiện tuần này</span>
                    <span class="col-week">${escapeHtml(currTitle)}</span>
                </div>
                ${renderItems(currWeek.section1)}
            </div>
        `;
    }

    function renderDetailComparison(prevWeek, currWeek) {
        const table = $('detailCompareTable');
        table.className = 'data-table detail-compare-table';
        const prevData = prevWeek.table_data || [];
        const currData = currWeek.table_data || [];

        // Build a map of units from both weeks
        const unitMap = new Map();
        prevData.forEach(r => unitMap.set(r.don_vi, { prev: r }));
        currData.forEach(r => {
            const existing = unitMap.get(r.don_vi) || {};
            existing.curr = r;
            unitMap.set(r.don_vi, existing);
        });

        const prevTitle = (prevWeek.title || 'Tuần trước').replace(/[()]/g, '').trim();
        const currTitle = (currWeek.title || 'Tuần này').replace(/[()]/g, '').trim();
        const shortPrev = prevTitle.length > 25 ? prevTitle.substring(0, 25) + '…' : prevTitle;
        const shortCurr = currTitle.length > 25 ? currTitle.substring(0, 25) + '…' : currTitle;

        let html = `<thead>
            <tr>
                <th rowspan="2">Đơn vị</th>
                <th colspan="3" class="th-group-prev">${escapeHtml(shortPrev)}</th>
                <th colspan="3" class="th-group-curr">${escapeHtml(shortCurr)}</th>
                <th colspan="2" class="th-group-diff">Chênh lệch</th>
            </tr>
            <tr>
                <th class="th-group-prev">TH tuần</th>
                <th class="th-group-prev">Lũy kế</th>
                <th class="th-group-prev">Còn lại</th>
                <th class="th-group-curr">TH tuần</th>
                <th class="th-group-curr">Lũy kế</th>
                <th class="th-group-curr">Còn lại</th>
                <th class="th-group-diff">TH tuần</th>
                <th class="th-group-diff">Lũy kế</th>
            </tr>
        </thead><tbody>`;

        for (const [unit, data] of unitMap) {
            const p = data.prev || { thuc_hien: 0, luy_ke: 0, con_lai: 0 };
            const c = data.curr || { thuc_hien: 0, luy_ke: 0, con_lai: 0 };
            const diffTH = c.thuc_hien - p.thuc_hien;
            const diffLK = c.luy_ke - p.luy_ke;
            const clsTH = diffTH > 0 ? 'diff-positive' : diffTH < 0 ? 'diff-negative' : 'diff-neutral';
            const clsLK = diffLK > 0 ? 'diff-positive' : diffLK < 0 ? 'diff-negative' : 'diff-neutral';
            const arrowTH = diffTH > 0 ? '↑' : diffTH < 0 ? '↓' : '—';
            const arrowLK = diffLK > 0 ? '↑' : diffLK < 0 ? '↓' : '—';

            html += `<tr>
                <td class="unit-name">${escapeHtml(unit)}</td>
                <td class="num">${p.thuc_hien.toLocaleString()}</td>
                <td class="num">${p.luy_ke.toLocaleString()}</td>
                <td class="num">${p.con_lai.toLocaleString()}</td>
                <td class="num">${c.thuc_hien.toLocaleString()}</td>
                <td class="num">${c.luy_ke.toLocaleString()}</td>
                <td class="num">${c.con_lai.toLocaleString()}</td>
                <td class="num ${clsTH}">${arrowTH} ${Math.abs(diffTH).toLocaleString()}</td>
                <td class="num ${clsLK}">${arrowLK} ${Math.abs(diffLK).toLocaleString()}</td>
            </tr>`;
        }

        // Totals row
        const totPrevTH = prevData.reduce((s, r) => s + r.thuc_hien, 0);
        const totPrevLK = prevData.reduce((s, r) => s + r.luy_ke, 0);
        const totPrevCL = prevData.reduce((s, r) => s + r.con_lai, 0);
        const totCurrTH = currData.reduce((s, r) => s + r.thuc_hien, 0);
        const totCurrLK = currData.reduce((s, r) => s + r.luy_ke, 0);
        const totCurrCL = currData.reduce((s, r) => s + r.con_lai, 0);
        const totDiffTH = totCurrTH - totPrevTH;
        const totDiffLK = totCurrLK - totPrevLK;
        const clsTotTH = totDiffTH > 0 ? 'diff-positive' : totDiffTH < 0 ? 'diff-negative' : 'diff-neutral';
        const clsTotLK = totDiffLK > 0 ? 'diff-positive' : totDiffLK < 0 ? 'diff-negative' : 'diff-neutral';

        html += `<tr style="font-weight:700; border-top:2px solid var(--border-strong);">
            <td class="unit-name">TỔNG CỘNG</td>
            <td class="num">${totPrevTH.toLocaleString()}</td>
            <td class="num">${totPrevLK.toLocaleString()}</td>
            <td class="num">${totPrevCL.toLocaleString()}</td>
            <td class="num">${totCurrTH.toLocaleString()}</td>
            <td class="num">${totCurrLK.toLocaleString()}</td>
            <td class="num">${totCurrCL.toLocaleString()}</td>
            <td class="num ${clsTotTH}">${totDiffTH > 0 ? '↑' : totDiffTH < 0 ? '↓' : '—'} ${Math.abs(totDiffTH).toLocaleString()}</td>
            <td class="num ${clsTotLK}">${totDiffLK > 0 ? '↑' : totDiffLK < 0 ? '↓' : '—'} ${Math.abs(totDiffLK).toLocaleString()}</td>
        </tr></tbody>`;

        table.innerHTML = html;
    }

    function renderChart(canvasId, label, labels, data, color) {
        const ctx = $(canvasId).getContext('2d');
        if (charts[canvasId]) charts[canvasId].destroy();
        charts[canvasId] = new Chart(ctx, {
            type: 'bar',
            data: { labels, datasets: [{ label: label + ' (phiếu)', data, backgroundColor: color + '40', borderColor: color, borderWidth: 2, borderRadius: 6, hoverBackgroundColor: color + '80' }] },
            options: {
                responsive: true,
                plugins: {
                    legend: { labels: { color: '#94a3b8', font: { family: 'Inter' } } },
                    tooltip: {
                        callbacks: {
                            label: ctx => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString()} phiếu`
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: { color: 'rgba(255,255,255,0.05)' },
                        ticks: {
                            color: '#64748b', font: { family: 'Inter' },
                            callback: v => v.toLocaleString()
                        },
                        title: { display: true, text: 'Số phiếu KSTV', color: '#64748b', font: { family: 'Inter', size: 11 } }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { color: '#64748b', font: { family: 'Inter', size: 10 }, maxRotation: 25 },
                        title: { display: true, text: 'Tuần báo cáo', color: '#64748b', font: { family: 'Inter', size: 11 } }
                    }
                },
                animation: {
                    onComplete: function() {
                        const chart = this;
                        const ctxDraw = chart.ctx;
                        ctxDraw.save();
                        ctxDraw.font = 'bold 12px Inter';
                        ctxDraw.fillStyle = '#e2e8f0';
                        ctxDraw.textAlign = 'center';
                        chart.data.datasets[0].data.forEach((val, i) => {
                            const meta = chart.getDatasetMeta(0).data[i];
                            ctxDraw.fillText(val.toLocaleString(), meta.x, meta.y - 8);
                        });
                        ctxDraw.restore();
                    }
                }
            }
        });
    }

    function renderDeltas(allData) {
        const deltaCards = $('deltaCards');
        if (allData.length < 2) { deltaCards.innerHTML = '<div style="color:var(--text-tertiary)">Cần ít nhất 2 tuần để so sánh</div>'; return; }
        const last = allData[allData.length - 1], prev = allData[allData.length - 2];
        const metrics = [
            { label: 'Phiếu KSTV thực hiện trong tuần', unit: 'phiếu', valFn: d => (d.table_data || []).reduce((s, r) => s + r.thuc_hien, 0) },
            { label: 'Phiếu KSTV lũy kế đến hiện tại', unit: 'phiếu', valFn: d => (d.table_data || []).reduce((s, r) => s + r.luy_ke, 0) },
            { label: 'Phiếu KSTV còn lại phải hoàn thành', unit: 'phiếu', valFn: d => (d.table_data || []).reduce((s, r) => s + r.con_lai, 0), invert: true },
        ];
        deltaCards.innerHTML = metrics.map(m => {
            const curr = m.valFn(last), pre = m.valFn(prev), diff = curr - pre;
            const pctChange = pre !== 0 ? ((diff / pre) * 100).toFixed(1) : (diff > 0 ? '∞' : '0');
            const isGood = m.invert ? diff <= 0 : diff >= 0;
            const cls = diff === 0 ? 'delta-neutral' : (isGood ? 'delta-up' : 'delta-down');
            const arrow = diff > 0 ? '↑' : diff < 0 ? '↓' : '—';
            return `<div class="delta-card">
                <div class="delta-label">${m.label}</div>
                <div class="delta-value ${cls}">${arrow} ${Math.abs(diff).toLocaleString()} ${m.unit}</div>
                <div class="delta-sub ${cls}">${diff >= 0 ? '+' : ''}${pctChange}% so với tuần trước</div>
                <div class="delta-detail">${pre.toLocaleString()} → ${curr.toLocaleString()} ${m.unit}</div>
            </div>`;
        }).join('');
    }

    // ─── VIEW MANAGEMENT ───
    function showView(view) {
        const driveView = $('driveView');
        const homeView = $('homeView');
        const presentationView = $('presentationView');

        if (reportView) reportView.classList.add('hidden');
        if (compareArea) compareArea.classList.add('hidden');
        if (emptyState) emptyState.classList.add('hidden');
        if (driveView) driveView.classList.add('hidden');
        if (homeView) homeView.classList.add('hidden');
        if (presentationView) presentationView.classList.add('hidden');

        // Update nav active styles
        $('btnSidebarHome')?.classList.remove('active-nav');
        $('btnSidebarPresentation')?.classList.remove('active-nav');
        $('btnOpenDriveSidebar')?.classList.remove('active-nav');

        if (view === 'home') {
            if (homeView) homeView.classList.remove('hidden');
            $('btnSidebarHome')?.classList.add('active-nav');
        } else if (view === 'presentation') {
            if (presentationView) presentationView.classList.remove('hidden');
            $('btnSidebarPresentation')?.classList.add('active-nav');
            initPresentationPlayer();
        } else if (view === 'report') {
            if (reportView) reportView.classList.remove('hidden');
        } else if (view === 'compare') {
            if (compareArea) compareArea.classList.remove('hidden');
        } else if (view === 'drive') {
            if (driveView) driveView.classList.remove('hidden');
            $('btnOpenDriveSidebar')?.classList.add('active-nav');
            renderDriveExplorer();
        } else {
            if (homeView) homeView.classList.remove('hidden');
            $('btnSidebarHome')?.classList.add('active-nav');
        }
    }
    window.appShowView = showView;

    function escapeHtml(s) {
        return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }

    // ─── REMINDER MODAL & GENERATOR ───
    const btnOpenReminder = $('btnOpenReminder');
    const closeReminderModal = $('closeReminderModal');
    const reminderModal = $('reminderModal');
    const btnCopyReminder = $('btnCopyReminder');
    const copyToast = $('copyToast');
    const reminderTextarea = $('reminderTextarea');

    function openReminderModal() {
        if (!activeComparisonData) {
            alert('Vui lòng chọn ít nhất 2 tuần để so sánh trước khi tạo thông báo!');
            return;
        }
        const text = generateReminderMessage(activeComparisonData.prevWeek, activeComparisonData.currWeek);
        if (reminderTextarea) reminderTextarea.value = text;
        if (reminderModal) reminderModal.classList.remove('hidden');
        if (copyToast) copyToast.classList.add('hidden');
    }

    if (btnOpenReminder) btnOpenReminder.addEventListener('click', openReminderModal);
    if (closeReminderModal) closeReminderModal.addEventListener('click', () => { if (reminderModal) reminderModal.classList.add('hidden'); });
    if (reminderModal) {
        reminderModal.addEventListener('click', (e) => {
            if (e.target === reminderModal) reminderModal.classList.add('hidden');
        });
    }

    if (btnCopyReminder && reminderTextarea) {
        btnCopyReminder.addEventListener('click', async () => {
            try {
                await navigator.clipboard.writeText(reminderTextarea.value);
                if (copyToast) copyToast.classList.remove('hidden');
                setTimeout(() => { if (copyToast) copyToast.classList.add('hidden'); }, 3000);
            } catch (err) {
                reminderTextarea.select();
                document.execCommand('copy');
                if (copyToast) copyToast.classList.remove('hidden');
                setTimeout(() => { if (copyToast) copyToast.classList.add('hidden'); }, 3000);
            }
        });
    }

    function generateReminderMessage(prevWeek, currWeek) {
        const currTitle = (currWeek.title || 'Tuần này').replace(/[()]/g, '').trim();
        const prevTitle = (prevWeek.title || 'Tuần trước').replace(/[()]/g, '').trim();

        const sumVal = (data, key) => (data.table_data || []).reduce((s, r) => s + r[key], 0);

        const currTH = sumVal(currWeek, 'thuc_hien');
        const prevTH = sumVal(prevWeek, 'thuc_hien');
        const diffTH = currTH - prevTH;
        const currLK = sumVal(currWeek, 'luy_ke');
        const currKH = sumVal(currWeek, 'ke_hoach');
        const currCL = sumVal(currWeek, 'con_lai');
        const pctProgress = currKH > 0 ? ((currLK / currKH) * 100).toFixed(1) : '0';

        const diffText = diffTH > 0 
            ? `TĂNG ${diffTH.toLocaleString()} phiếu (+${prevTH ? ((diffTH/prevTH)*100).toFixed(1) : 0}%)`
            : diffTH < 0 
                ? `GIẢM ${Math.abs(diffTH).toLocaleString()} phiếu (${((diffTH/prevTH)*100).toFixed(1)}%)`
                : 'Giữ nguyên số lượng';

        // 1. Slow/zero units
        const currData = currWeek.table_data || [];
        const slowUnits = [];
        currData.forEach(r => {
            if (r.thuc_hien === 0) {
                slowUnits.push(`• Đơn vị [${r.don_vi}]: 0 phiếu trong tuần (Còn lại ${r.con_lai.toLocaleString()} phiếu) -> Yêu cầu khẩn trương triển khai.`);
            } else if (r.thuc_hien < 10) {
                slowUnits.push(`• Đơn vị [${r.don_vi}]: Hoàn thành ${r.thuc_hien} phiếu trong tuần (Còn lại ${r.con_lai.toLocaleString()} phiếu) -> Cần tiếp tục tăng tốc.`);
            }
        });

        // 2. Unfinished tasks from prev week plan
        const planItems = (prevWeek.section2 || []).filter(i => !i.bold && i.text && i.text.trim().length >= 5 && !isIgnoredNonTaskLine(i.text) && !i.text.trim().startsWith('+'));
        const resultTexts = (currWeek.section1 || []).filter(r => !r.bold).map(r => r.text.toLowerCase()).join(' ');

        const stopWords = new Set(['và','các','của','cho','với','trong','từ','về','theo','đã','được','có','để','tại','do','là','một','này','đó','trên','khi','sẽ','đến','ra','lên','không','những','công','tác','việc','tuần','ngày','năm','tháng']);
        function extractKeywords(text) {
            return text.toLowerCase().replace(/[.,;:!?()"']/g, ' ').split(/\s+/).filter(w => w.length > 2 && !stopWords.has(w) && !/^\d+$/.test(w));
        }

        const pendingTasks = [];
        planItems.forEach(item => {
            const kws = extractKeywords(item.text);
            if (kws.length === 0) return;
            let matches = 0;
            kws.forEach(kw => { if (resultTexts.includes(kw)) matches++; });
            const score = matches / kws.length;
            if (score < 0.5) {
                pendingTasks.push(`• ${item.text.trim()}`);
            }
        });

        // 3. Current week recommendations
        const recItems = (currWeek.section3 || []).filter(i => !i.bold && i.text && i.text.trim().length >= 5 && !isIgnoredNonTaskLine(i.text));

        let text = `📢 [THÔNG BÁO / ĐÔN ĐỐC CÔNG TÁC ATVSLĐ GỬI KỸ SƯ AN TOÀN (KSAT)]\n`;
        text += `🗓 Báo cáo tuần: ${currTitle}\n`;
        text += `(Đối chiếu kế hoạch tuần trước: ${prevTitle})\n`;
        text += `--------------------------------------------------\n\n`;

        text += `1. 📊 TỔNG QUAN KẾT QUẢ KSTV:\n`;
        text += `• Thực hiện trong tuần: ${currTH.toLocaleString()} phiếu (${diffText} so với tuần trước).\n`;
        text += `• Tiến độ lũy kế: ${currLK.toLocaleString()} / ${currKH.toLocaleString()} phiếu (Đạt ${pctProgress}% kế hoạch năm).\n`;
        text += `• Khối lượng còn lại: ${currCL.toLocaleString()} phiếu.\n\n`;

        text += `2. ⚠️ ĐÔN ĐỐC TIẾN ĐỘ TẠI CÁC ĐƠN VỊ:\n`;
        if (slowUnits.length > 0) {
            text += slowUnits.join('\n') + `\n\n`;
        } else {
            text += `• Tiến độ thực hiện các đơn vị tương đối tốt, yêu cầu tiếp tục duy trì và đẩy nhanh sản lượng.\n\n`;
        }

        text += `3. 🚨 NHIỆM VỤ TỒN ĐỌNG CHƯA HOÀN THÀNH TỪ KẾ HOẠCH TUẦN TRƯỚC:\n`;
        if (pendingTasks.length > 0) {
            text += pendingTasks.join('\n') + `\n\n`;
        } else {
            text += `• Các mục tiêu kế hoạch tuần trước đã hoàn thành tốt.\n\n`;
        }

        text += `4. 🛡 CHÚ Ý & YÊU CẦU ĐỐI VỚI LỰC LƯỢNG KỸ SƯ AN TOÀN (KSAT):\n`;
        if (recItems.length > 0) {
            recItems.forEach(r => { text += `• ${r.text.trim()}\n`; });
        }
        text += `• Tăng cường kiểm tra, giám sát trực tiếp hiện trường công tác, đảm bảo an toàn tuyệt đối.\n`;
        text += `• Đôn đốc giải quyết triệt để các tồn đọng về kiểm tra ATLĐ và hoàn thành chỉ tiêu phiếu KSTV được giao.\n\n`;

        text += `--------------------------------------------------\n`;
        text += `(Văn bản tự động tổng hợp từ Hệ thống Dashboard Quản lý Báo cáo KSAT)`;

        return text;
    }

    // ═══════════════════════════════════════════════════════════
    //  SHARED DRIVE STORE & EXPLORER CONTROLLER
    // ═══════════════════════════════════════════════════════════
    const DEFAULT_DRIVE_URL = 'https://drive.google.com/drive/folders/17_809JzLyZoYKDP2JDoPhwwTn912HCKQ?usp=sharing';
    
    const initialDriveItems = [
        {
            id: 'f1', parentId: 'root', type: 'folder', name: 'Báo cáo Công tác An toàn & Phiếu KSTV',
            theme: 'blue', desc: 'Báo cáo tuần, phiếu KSTV điện gia đình các đơn vị', date: '2026-07-20T08:00:00.000Z'
        },
        {
            id: 'f2', parentId: 'root', type: 'folder', name: 'Văn bản Chỉ đạo & Quy trình ATVSLĐ',
            theme: 'purple', desc: 'Quy định EVN, PCVT, Nghị định 44, Thông tư ATLĐ', date: '2026-07-15T08:00:00.000Z'
        },
        {
            id: 'f3', parentId: 'root', type: 'folder', name: 'Biểu mẫu & Sổ theo dõi KSAT',
            theme: 'amber', desc: 'Mẫu biên bản kiểm tra, sổ theo dõi cấp phát trang cụ an toàn', date: '2026-07-10T08:00:00.000Z'
        },
        {
            id: 'f4', parentId: 'root', type: 'folder', name: 'Hình ảnh & Video Giám sát Hiện trường',
            theme: 'green', desc: 'Ảnh kiểm tra Hotline, tỉa cây xanh, bảo dưỡng rơ-le', date: '2026-07-18T08:00:00.000Z'
        },
        {
            id: 'f5', parentId: 'root', type: 'folder', name: 'Thử nghiệm & Dụng cụ An toàn (ETC)',
            theme: 'cyan', desc: 'Biên bản thử nghiệm định kỳ găng tay, sào, dây an toàn', date: '2026-07-12T08:00:00.000Z'
        }
    ];

    function getLocalDriveItems() {
        const saved = localStorage.getItem('ksat_drive_items');
        if (saved) {
            try {
                let items = JSON.parse(saved);
                // Filter out old sample file items
                items = items.filter(i => !['item1', 'item2', 'item3', 'item4'].includes(i.id));
                localStorage.setItem('ksat_drive_items', JSON.stringify(items));
                return items;
            } catch(e){}
        }
        localStorage.setItem('ksat_drive_items', JSON.stringify(initialDriveItems));
        return initialDriveItems;
    }

    async function syncDriveItemsOnline() {
        const online = await store.getDriveItems();
        if (online && online.length) {
            const map = new Map();
            initialDriveItems.forEach(i => map.set(i.id, i));
            online.forEach(i => map.set(i.id, i));
            const merged = Array.from(map.values());
            localStorage.setItem('ksat_drive_items', JSON.stringify(merged));
            renderDriveExplorer();
        }
    }

    function saveDriveItems(items) {
        localStorage.setItem('ksat_drive_items', JSON.stringify(items));
    }

    let driveCurrentFolderId = 'root';
    let driveActiveFilter = 'all';
    let driveSearchQuery = '';

    const btnOpenDriveSidebar = $('btnOpenDriveSidebar');
    const btnEmptyDrive = $('btnEmptyDrive');
    const tabDriveBtn = $('tabDriveBtn');

    if (btnOpenDriveSidebar) btnOpenDriveSidebar.addEventListener('click', () => { closeSidebar(); showView('drive'); });
    if (btnEmptyDrive) btnEmptyDrive.addEventListener('click', () => showView('drive'));
    if (tabDriveBtn) tabDriveBtn.addEventListener('click', () => showView('drive'));

    function renderDriveExplorer() {
        const driveGrid = $('driveGrid');
        const driveBreadcrumb = $('driveBreadcrumb');
        const driveStatsBadge = $('driveStatsBadge');
        if (!driveGrid) return;

        const allItems = getLocalDriveItems();

        // 1. Breadcrumbs build
        const folderMap = new Map(allItems.filter(i => i.type === 'folder').map(f => [f.id, f]));
        const crumbs = [];
        let currId = driveCurrentFolderId;
        while (currId && currId !== 'root') {
            const f = folderMap.get(currId);
            if (f) {
                crumbs.unshift({ id: f.id, name: f.name });
                currId = f.parentId;
            } else { break; }
        }
        crumbs.unshift({ id: 'root', name: '🏠 Thư Mục Gốc' });

        driveBreadcrumb.innerHTML = crumbs.map((c, idx) => {
            const isLast = idx === crumbs.length - 1;
            return `<span class="crumb-item ${isLast ? 'active' : ''}" onclick="window.navDriveFolder('${c.id}')">${escapeHtml(c.name)}</span>
            ${!isLast ? '<span class="crumb-sep">/</span>' : ''}`;
        }).join('');

        // 2. Filter items in current folder or by search
        let filtered = allItems.filter(i => {
            if (driveSearchQuery) {
                return i.name.toLowerCase().includes(driveSearchQuery.toLowerCase()) ||
                       (i.desc && i.desc.toLowerCase().includes(driveSearchQuery.toLowerCase()));
            }
            return i.parentId === driveCurrentFolderId;
        });

        if (driveActiveFilter !== 'all') {
            filtered = filtered.filter(i => {
                if (driveActiveFilter === 'folder') return i.type === 'folder';
                if (driveActiveFilter === 'docx') return i.type === 'docx';
                if (driveActiveFilter === 'xlsx') return i.type === 'xlsx';
                if (driveActiveFilter === 'pdf') return i.type === 'pdf';
                if (driveActiveFilter === 'media') return i.type === 'media';
                if (driveActiveFilter === 'link') return i.type === 'link';
                return true;
            });
        }

        driveStatsBadge.textContent = `${filtered.length} mục`;

        if (!filtered.length) {
            driveGrid.innerHTML = `
                <div class="empty-drive-msg">
                    <p style="font-size:1.1rem;margin-bottom:0.5rem">📂 Thư mục này chưa có dữ liệu</p>
                    <p style="font-size:0.8rem;color:var(--text-tertiary)">Bấm "Tạo Thư Mục" hoặc "Thêm File / Link" ở trên để tải dữ liệu lên kho dùng chung.</p>
                </div>`;
            return;
        }

        // Sort folders first, then files
        filtered.sort((a, b) => {
            if (a.type === 'folder' && b.type !== 'folder') return -1;
            if (a.type !== 'folder' && b.type === 'folder') return 1;
            return a.name.localeCompare(b.name, 'vi');
        });

        driveGrid.innerHTML = filtered.map(item => {
            if (item.type === 'folder') {
                const subCount = allItems.filter(i => i.parentId === item.id).length;
                const themeClass = item.theme ? `icon-${item.theme}` : 'icon-blue';
                return `
                    <div class="drive-item-card" onclick="window.navDriveFolder('${item.id}')">
                        <div class="drive-card-top">
                            <div class="drive-icon-badge ${themeClass}">📂</div>
                            <span class="drive-item-type-tag badge-folder">Thư mục</span>
                        </div>
                        <div class="drive-card-body">
                            <div class="drive-item-title">${escapeHtml(item.name)}</div>
                            <div class="drive-item-desc">${escapeHtml(item.desc || 'Chứa các văn bản, tài liệu liên quan')}</div>
                        </div>
                        <div class="drive-card-footer">
                            <span>${subCount} mục bên trong</span>
                            <div class="drive-action-links" onclick="event.stopPropagation()">
                                <button class="drive-btn-icon drive-btn-delete" onclick="window.deleteDriveItem('${item.id}')" title="Xóa thư mục">🗑</button>
                            </div>
                        </div>
                    </div>`;
            } else {
                const badgeMap = {
                    docx: { icon: '📄', class: 'badge-docx', label: 'Word' },
                    xlsx: { icon: '📊', class: 'badge-xlsx', label: 'Excel' },
                    pdf: { icon: '📕', class: 'badge-pdf', label: 'PDF' },
                    media: { icon: '🖼️', class: 'badge-media', label: 'Media' },
                    link: { icon: '🔗', class: 'badge-link', label: 'Drive Link' },
                    note: { icon: '📝', class: 'badge-note', label: 'Ghi chú' }
                };
                const meta = badgeMap[item.type] || badgeMap.link;
                const fileUrl = item.url || '';
                const openAction = fileUrl ? `onclick="window.open('${escapeHtml(fileUrl)}', '_blank')"` : '';

                const openLinkHtml = fileUrl ? `
                    <a href="${escapeHtml(fileUrl)}" target="_blank" download="${escapeHtml(item.name)}" rel="noopener noreferrer" class="drive-btn-icon" title="Mở hoặc tải file">
                        <span>🚀 Mở File</span>
                    </a>` : `<span class="drive-btn-icon" style="opacity:0.6;cursor:default">📄 Tệp lưu trữ</span>`;

                return `
                    <div class="drive-item-card">
                        <div class="drive-card-top">
                            <div class="drive-icon-badge icon-blue">${meta.icon}</div>
                            <span class="drive-item-type-tag ${meta.class}">${meta.label}</span>
                        </div>
                        <div class="drive-card-body">
                            <div class="drive-item-title" ${openAction}>${escapeHtml(item.name)}</div>
                            <div class="drive-item-desc">${escapeHtml(item.desc || 'Tài liệu dùng chung')}</div>
                        </div>
                        <div class="drive-card-footer">
                            <span>${item.date ? new Date(item.date).toLocaleDateString('vi-VN') : 'Gần đây'}</span>
                            <div class="drive-action-links">
                                ${openLinkHtml}
                                <button class="drive-btn-icon drive-btn-delete" onclick="window.deleteDriveItem('${item.id}')" title="Xóa tài liệu">🗑</button>
                            </div>
                        </div>
                    </div>`;
            }
        }).join('');
    }

    window.navDriveFolder = function(folderId) {
        driveCurrentFolderId = folderId;
        renderDriveExplorer();
    };

    window.deleteDriveItem = function(itemId) {
        if (!confirm('Bạn có chắc muốn xóa mục này khỏi thư mục dùng chung?')) return;
        let items = getLocalDriveItems();
        items = items.filter(i => i.id !== itemId && i.parentId !== itemId);
        saveDriveItems(items);
        store.removeDriveItem(itemId);
        renderDriveExplorer();
    };

    // Filter pills event
    const filterContainer = $('driveFilterPills');
    if (filterContainer) {
        filterContainer.querySelectorAll('.filter-pill').forEach(btn => {
            btn.addEventListener('click', () => {
                filterContainer.querySelectorAll('.filter-pill').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                driveActiveFilter = btn.dataset.filter;
                renderDriveExplorer();
            });
        });
    }

    // Search input event
    const driveSearchInput = $('driveSearchInput');
    if (driveSearchInput) {
        driveSearchInput.addEventListener('input', e => {
            driveSearchQuery = e.target.value.trim();
            renderDriveExplorer();
        });
    }

    // Modal Create Folder
    const btnCreateFolder = $('btnCreateFolder');
    const createFolderModal = $('createFolderModal');
    const closeCreateFolderModal = $('closeCreateFolderModal');
    const btnCancelFolder = $('btnCancelFolder');
    const btnSaveFolder = $('btnSaveFolder');

    if (btnCreateFolder) btnCreateFolder.addEventListener('click', () => {
        if (createFolderModal) createFolderModal.classList.remove('hidden');
    });
    const hideFolderModal = () => { if (createFolderModal) createFolderModal.classList.add('hidden'); };
    if (closeCreateFolderModal) closeCreateFolderModal.addEventListener('click', hideFolderModal);
    if (btnCancelFolder) btnCancelFolder.addEventListener('click', hideFolderModal);

    if (btnSaveFolder) {
        btnSaveFolder.addEventListener('click', () => {
            const name = $('folderNameInput').value.trim();
            const theme = $('folderThemeSelect').value;
            const desc = $('folderDescInput').value.trim();

            if (!name) { alert('Vui lòng nhập tên thư mục!'); return; }

            const newFolder = {
                id: 'f_' + Date.now(),
                parentId: driveCurrentFolderId,
                type: 'folder',
                name, theme, desc,
                date: new Date().toISOString()
            };

            const items = getLocalDriveItems();
            items.push(newFolder);
            saveDriveItems(items);
            store.saveDriveItem(newFolder);

            $('folderNameInput').value = '';
            $('folderDescInput').value = '';
            hideFolderModal();
            renderDriveExplorer();
        });
    }

    // Modal Add File / Link
    const btnAddFile = $('btnAddFile');
    const addFileModal = $('addFileModal');
    const closeAddFileModal = $('closeAddFileModal');
    const btnCancelFile = $('btnCancelFile');
    const btnSaveFile = $('btnSaveFile');
    const localFileInput = $('localFileInput');
    let pendingFileContent = null;

    if (localFileInput) {
        localFileInput.addEventListener('change', e => {
            const file = e.target.files[0];
            if (!file) return;

            const fileNameInput = $('fileNameInput');
            if (fileNameInput && !fileNameInput.value.trim()) {
                fileNameInput.value = file.name;
            }

            const ext = file.name.split('.').pop().toLowerCase();
            const fileTypeSelect = $('fileTypeSelect');
            if (fileTypeSelect) {
                if (['docx', 'doc'].includes(ext)) fileTypeSelect.value = 'docx';
                else if (['xlsx', 'xls', 'csv'].includes(ext)) fileTypeSelect.value = 'xlsx';
                else if (ext === 'pdf') fileTypeSelect.value = 'pdf';
                else if (['jpg', 'jpeg', 'png', 'gif', 'mp4', 'mov'].includes(ext)) fileTypeSelect.value = 'media';
                else fileTypeSelect.value = 'note';
            }

            if (file.size <= 10 * 1024 * 1024) {
                const reader = new FileReader();
                reader.onload = ev => { pendingFileContent = ev.target.result; };
                reader.readAsDataURL(file);
            } else {
                pendingFileContent = null;
            }
        });
    }

    if (btnAddFile) btnAddFile.addEventListener('click', () => {
        if (addFileModal) addFileModal.classList.remove('hidden');
    });
    const hideFileModal = () => {
        if (addFileModal) addFileModal.classList.add('hidden');
        pendingFileContent = null;
        if (localFileInput) localFileInput.value = '';
    };
    if (closeAddFileModal) closeAddFileModal.addEventListener('click', hideFileModal);
    if (btnCancelFile) btnCancelFile.addEventListener('click', hideFileModal);

    function readFileAsDataURL(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = ev => resolve(ev.target.result);
            reader.onerror = err => reject(err);
            reader.readAsDataURL(file);
        });
    }

    if (btnSaveFile) {
        btnSaveFile.addEventListener('click', async () => {
            const name = $('fileNameInput').value.trim();
            const urlInput = $('fileUrlInput').value.trim();
            const type = $('fileTypeSelect').value;
            const desc = $('fileDescInput').value.trim();

            if (!name) { alert('Vui lòng chọn file hoặc nhập tên tài liệu!'); return; }

            btnSaveFile.disabled = true;
            btnSaveFile.textContent = '⏳ Đang tải file lên Google Drive...';

            let fileUrl = urlInput || '';
            const selectedFile = localFileInput && localFileInput.files && localFileInput.files[0];
            if (selectedFile) {
                try {
                    fileUrl = await readFileAsDataURL(selectedFile);
                } catch (e) {
                    console.error('Lỗi đọc file:', e);
                }
            }

            const newFile = {
                id: 'file_' + Date.now(),
                parentId: driveCurrentFolderId,
                type, name,
                url: fileUrl,
                desc: desc || (selectedFile ? `Tệp từ máy tính (${selectedFile.name})` : 'Tài liệu dùng chung'),
                date: new Date().toISOString()
            };

            const items = getLocalDriveItems();
            items.push(newFile);
            saveDriveItems(items);

            try {
                const res = await store.saveDriveItem(newFile);
                if (res && res.item && res.item.url) {
                    newFile.url = res.item.url;
                    saveDriveItems(items);
                }
            } catch(e) {}

            btnSaveFile.disabled = false;
            btnSaveFile.textContent = 'Thêm Tài Liệu';

            $('fileNameInput').value = '';
            $('fileUrlInput').value = '';
            $('fileDescInput').value = '';
            hideFileModal();
            renderDriveExplorer();
        });
    }

    // Desktop Sidebar Toggle
    const btnToggleSidebarDesktop = $('btnToggleSidebarDesktop');
    if (btnToggleSidebarDesktop && sidebar) {
        btnToggleSidebarDesktop.addEventListener('click', () => {
            sidebar.classList.toggle('collapsed');
            const isCollapsed = sidebar.classList.contains('collapsed');
            btnToggleSidebarDesktop.innerHTML = isCollapsed
                ? `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="18" height="18"><path d="M9 18l6-6-6-6"/></svg>`
                : `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="18" height="18"><path d="M15 18l-6-6 6-6"/></svg>`;
        });
    }

    const btnSidebarHome = $('btnSidebarHome');
    const btnSidebarPresentation = $('btnSidebarPresentation');

    if (btnSidebarHome) btnSidebarHome.addEventListener('click', () => { closeSidebar(); showView('home'); });
    if (btnSidebarPresentation) btnSidebarPresentation.addEventListener('click', () => { closeSidebar(); showView('presentation'); });

    // ═══════════════════════════════════════════════════════════
    //  SLIDE PRESENTATION PLAYER (Sinh hoạt an toàn)
    // ═══════════════════════════════════════════════════════════
    let currentSlideIndex = 0;
    const TOTAL_SLIDES = 26;
    let slideAutoInterval = null;

    function renderSlide(index) {
        if (index < 0) index = 0;
        if (index >= TOTAL_SLIDES) index = TOTAL_SLIDES - 1;
        currentSlideIndex = index;

        const currentSlideImg = $('currentSlideImg');
        const slideNumberBadge = $('slideNumberBadge');
        const slideInputNumber = $('slideInputNumber');

        if (currentSlideImg) {
            currentSlideImg.style.opacity = '0.4';
            currentSlideImg.src = `slides/2026-07-slides/Slide${index + 1}.JPG`;
            setTimeout(() => { currentSlideImg.style.opacity = '1'; }, 80);
        }
        if (slideNumberBadge) slideNumberBadge.textContent = `Slide ${index + 1} / ${TOTAL_SLIDES}`;
        if (slideInputNumber) slideInputNumber.value = index + 1;

        const thumbsScroll = $('slideThumbsScroll');
        if (thumbsScroll) {
            thumbsScroll.querySelectorAll('.slide-thumb-item').forEach((item, idx) => {
                if (idx === index) {
                    item.classList.add('active-thumb');
                    item.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
                } else {
                    item.classList.remove('active-thumb');
                }
            });
        }
    }

    function initPresentationPlayer() {
        const thumbsScroll = $('slideThumbsScroll');
        if (thumbsScroll && thumbsScroll.children.length === 0) {
            let html = '';
            for (let i = 1; i <= TOTAL_SLIDES; i++) {
                html += `
                    <div class="slide-thumb-item ${i === 1 ? 'active-thumb' : ''}" onclick="window.appRenderSlide(${i - 1})">
                        <img src="slides/2026-07-slides/Slide${i}.JPG" alt="Slide ${i}" loading="lazy">
                        <span>Slide ${i}</span>
                    </div>`;
            }
            thumbsScroll.innerHTML = html;
        }
        renderSlide(currentSlideIndex);
    }
    window.appRenderSlide = renderSlide;

    const btnPrevSlide = $('btnPrevSlide');
    const btnNextSlide = $('btnNextSlide');
    const slideInputNumber = $('slideInputNumber');
    const btnFullscreenSlide = $('btnFullscreenSlide');
    const btnToggleAutoPlay = $('btnToggleAutoPlay');
    const autoPlayLabel = $('autoPlayLabel');
    const btnToggleThumbnails = $('btnToggleThumbnails');
    const btnCloseThumbs = $('btnCloseThumbs');
    const slideThumbsContainer = $('slideThumbsContainer');

    if (btnPrevSlide) btnPrevSlide.addEventListener('click', () => renderSlide(currentSlideIndex - 1));
    if (btnNextSlide) btnNextSlide.addEventListener('click', () => renderSlide(currentSlideIndex + 1));
    if (slideInputNumber) slideInputNumber.addEventListener('change', e => {
        const val = parseInt(e.target.value);
        if (!isNaN(val)) renderSlide(val - 1);
    });

    if (btnFullscreenSlide) {
        btnFullscreenSlide.addEventListener('click', () => {
            const stage = $('presentationStage');
            if (stage) {
                if (!document.fullscreenElement) {
                    stage.requestFullscreen().catch(err => alert('Không thể mở toàn màn hình: ' + err.message));
                } else {
                    document.exitFullscreen();
                }
            }
        });
    }

    if (btnToggleAutoPlay) {
        btnToggleAutoPlay.addEventListener('click', () => {
            if (slideAutoInterval) {
                clearInterval(slideAutoInterval);
                slideAutoInterval = null;
                if (autoPlayLabel) autoPlayLabel.textContent = '▶ Tự động';
                btnToggleAutoPlay.classList.remove('primary-btn');
            } else {
                slideAutoInterval = setInterval(() => {
                    let next = currentSlideIndex + 1;
                    if (next >= TOTAL_SLIDES) next = 0;
                    renderSlide(next);
                }, 4000);
                if (autoPlayLabel) autoPlayLabel.textContent = '⏸ Tạm dừng';
                btnToggleAutoPlay.classList.add('primary-btn');
            }
        });
    }

    if (btnToggleThumbnails && slideThumbsContainer) {
        btnToggleThumbnails.addEventListener('click', () => {
            slideThumbsContainer.classList.toggle('hidden');
        });
    }
    if (btnCloseThumbs && slideThumbsContainer) {
        btnCloseThumbs.addEventListener('click', () => {
            slideThumbsContainer.classList.add('hidden');
        });
    }

    document.addEventListener('keydown', e => {
        const presentationView = $('presentationView');
        if (presentationView && !presentationView.classList.contains('hidden')) {
            if (e.key === 'ArrowLeft') renderSlide(currentSlideIndex - 1);
            else if (e.key === 'ArrowRight') renderSlide(currentSlideIndex + 1);
        }
    });

    // Default view: Home
    showView('home');

    // Initial background sync for shared drive items
    syncDriveItemsOnline();
});
