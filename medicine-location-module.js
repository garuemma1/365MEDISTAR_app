/**
 * 📦 365메디스타약국 일반약 보관 위치 관리 모듈 (Medicine Location Module v7.0)
 * - 번호식 매대 구역 (1번 구역 ~ 12번 구역 + 카운터/냉장/창고 + 구역 자유 추가)
 * - 초기 엑셀 품목 DB (9,610품목) 기본 탑재 + 추가 엑셀/CSV 일괄 업로드 지원
 * - 코페이 POS(글자 전용) 보완용 현장 다중 사진(최대 5장) + 초성 검색 + 위치 이동 타임라인 추적
 */
window.MedicineLocationModule = (function () {
  let activeCategory = 'ALL';
  let searchQuery = '';
  let excelPageLimit = 60;

  function escapeHTML(str) {
    if (!str) return '';
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  }

  function getInitialSound(text) {
    if (!text) return '';
    const CHOSUNG = [
      'ㄱ', 'ㄲ', 'ㄴ', 'ㄷ', 'ㄸ', 'ㄹ', 'ㅁ', 'ㅂ', 'ㅃ',
      'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅉ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ'
    ];
    let result = '';
    for (let i = 0; i < text.length; i++) {
      const code = text.charCodeAt(i) - 44032;
      if (code >= 0 && code <= 11172) {
        const chosungIndex = Math.floor(code / 588);
        result += CHOSUNG[chosungIndex];
      } else {
        result += text.charAt(i);
      }
    }
    return result;
  }

  // 번호식 기본 매대 구역 (1번 ~ 12번 + 카운터/냉장고/창고)
  const BASE_DEFAULT_ZONES = [
    { id: 'ZONE_1', name: '1번 구역', icon: 'fa-hashtag', color: '#2563eb', bg: '#eff6ff', border: '#bfdbfe' },
    { id: 'ZONE_2', name: '2번 구역', icon: 'fa-hashtag', color: '#059669', bg: '#f0fdf4', border: '#86efac' },
    { id: 'ZONE_3', name: '3번 구역', icon: 'fa-hashtag', color: '#d97706', bg: '#fffbeb', border: '#fde68a' },
    { id: 'ZONE_4', name: '4번 구역', icon: 'fa-hashtag', color: '#7c3aed', bg: '#faf5ff', border: '#ddd6fe' },
    { id: 'ZONE_5', name: '5번 구역', icon: 'fa-hashtag', color: '#0891b2', bg: '#ecfeff', border: '#a5f3fc' },
    { id: 'ZONE_6', name: '6번 구역', icon: 'fa-hashtag', color: '#ea580c', bg: '#fff7ed', border: '#fed7aa' },
    { id: 'ZONE_7', name: '7번 구역', icon: 'fa-hashtag', color: '#4f46e5', bg: '#eef2ff', border: '#c7d2fe' },
    { id: 'ZONE_8', name: '8번 구역', icon: 'fa-hashtag', color: '#0d9488', bg: '#f0fdfa', border: '#99f6e4' },
    { id: 'ZONE_9', name: '9번 구역', icon: 'fa-hashtag', color: '#db2777', bg: '#fdf2f8', border: '#fbcfe8' },
    { id: 'ZONE_10', name: '10번 구역', icon: 'fa-hashtag', color: '#65a30d', bg: '#f7fee7', border: '#d9f99d' },
    { id: 'ZONE_11', name: '11번 구역', icon: 'fa-hashtag', color: '#2563eb', bg: '#eff6ff', border: '#bfdbfe' },
    { id: 'ZONE_12', name: '12번 구역', icon: 'fa-hashtag', color: '#059669', bg: '#f0fdf4', border: '#86efac' },
    { id: 'ZONE_COUNTER', name: '카운터 매대', icon: 'fa-cash-register', color: '#1d4ed8', bg: '#dbeafe', border: '#93c5fd' },
    { id: 'ZONE_REF', name: '❄️ 냉장고', icon: 'fa-snowflake', color: '#0284c7', bg: '#f0f9ff', border: '#bae6fd' },
    { id: 'ZONE_STORAGE', name: '📦 창고/여분박스', icon: 'fa-boxes-stacked', color: '#dc2626', bg: '#fff1f2', border: '#fecdd3' }
  ];

  function getZones() {
    try {
      const customRaw = localStorage.getItem('mds_custom_med_zones_v1');
      if (customRaw) {
        const parsed = JSON.parse(customRaw);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch (e) {}
    return BASE_DEFAULT_ZONES;
  }

  function addNextNumberZone() {
    const zones = [...getZones()];
    const numZones = zones.filter(z => /^ZONE_\d+$/.test(z.id));
    const nextNum = numZones.length + 1;
    const customLabel = prompt(`새로 추가할 매대 구역 이름을 입력하세요 (기본값: ${nextNum}번 구역):`, `${nextNum}번 구역`);
    if (!customLabel || !customLabel.trim()) return;
    const newId = 'ZONE_' + Date.now();
    zones.splice(numZones.length, 0, {
      id: newId,
      name: customLabel.trim(),
      icon: 'fa-hashtag',
      color: '#2563eb',
      bg: '#eff6ff',
      border: '#bfdbfe'
    });
    try {
      localStorage.setItem('mds_custom_med_zones_v1', JSON.stringify(zones));
    } catch (e) {}
    render('module-content');
  }

  // 초기 엑셀 품목 DB + 사용자 추가 엑셀 품목 통합 조회
  function getExcelCatalog() {
    const base = Array.isArray(window.OTC_EXCEL_PRESET_DB) ? window.OTC_EXCEL_PRESET_DB : [];
    let extra = [];
    try {
      const raw = localStorage.getItem('mds_custom_excel_otc_v1');
      if (raw) extra = JSON.parse(raw) || [];
    } catch (e) {}
    return extra.length > 0 ? [...extra, ...base] : base;
  }

  // 🕒 한국 로컬 일시(KST) 안전 포맷팅 헬퍼 함수
  function formatLocationDate(val, fallbackStr) {
    if (fallbackStr && typeof fallbackStr === 'string' && fallbackStr.length >= 10) {
      return fallbackStr.replace(/-/g, '.');
    }
    if (!val) return '';
    if (typeof val === 'number' || (!isNaN(Number(val)) && String(val).length >= 10 && !String(val).includes('-') && !String(val).includes('.'))) {
      const d = new Date(Number(val));
      if (!isNaN(d.getTime())) {
        const yy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        const hh = String(d.getHours()).padStart(2, '0');
        const mi = String(d.getMinutes()).padStart(2, '0');
        return `${yy}.${mm}.${dd} ${hh}:${mi}`;
      }
    }
    if (typeof val === 'string') {
      if (val.includes('-') || val.includes('.')) {
        return val.replace(/-/g, '.').substring(0, 16);
      }
    }
    return '';
  }

  function getStorageData() {
    try {
      if (window.SheetsSync && typeof window.SheetsSync.getMedicineLocations === 'function') {
        return window.SheetsSync.getMedicineLocations();
      }
      const raw = localStorage.getItem('mds_medicine_locations_v1') || localStorage.getItem('mds_medicine_locations_v1');
      return raw ? JSON.parse(raw) : [];
    } catch (e) {
      return [];
    }
  }

  function saveStorageData(list) {
    try {
      const json = JSON.stringify(list || []);
      localStorage.setItem('mds_medicine_locations_v1', json);
    } catch (e) {}

    if (window.SheetsSync && typeof window.SheetsSync.saveMedicineLocations === 'function') {
      window.SheetsSync.saveMedicineLocations(list);
      if (typeof window.SheetsSync.pushToCloud === 'function') {
        window.SheetsSync.pushToCloud();
      }
    }
  }

  // ⚡ 새로 등록/수정된 약품 위치 정보가 무조건 맨 최상단(Top) 첫 번째에 뜨도록 최신순 내림차순 정렬
  function sortItems(list) {
    return [...list].sort((a, b) => {
      const getNum = (item) => {
        if (item.updatedAt) {
          if (typeof item.updatedAt === 'number') return item.updatedAt;
          const parsed = Date.parse(String(item.updatedAt).replace(/\./g, '-'));
          if (!isNaN(parsed)) return parsed;
        }
        if (item.createdAt) {
          if (typeof item.createdAt === 'number') return item.createdAt;
          const parsed = Date.parse(String(item.createdAt).replace(/\./g, '-'));
          if (!isNaN(parsed)) return parsed;
        }
        if (item.id && typeof item.id === 'string' && item.id.startsWith('med_')) {
          const num = parseInt(item.id.replace('med_', ''), 10);
          if (!isNaN(num)) return num;
        }
        return 0;
      };
      return getNum(b) - getNum(a);
    });
  }

  function filterRegisteredItems(items, cat, queryStr) {
    if (cat === 'EXCEL_DB') return [];
    const rawQ = (queryStr || '').toLowerCase().trim();
    const keywords = rawQ.split(/\s+/).filter(Boolean);

    return items.filter(item => {
      const matchCat = cat === 'ALL' || item.zoneId === cat;
      if (!rawQ) return matchCat;

      const targetText = [
        item.name || '',
        item.zoneName || '',
        item.locationDetail || '',
        item.posLocationMemo || '',
        item.barcode || '',
        item.notes || '',
        item.updatedBy || ''
      ].join(' ').toLowerCase();

      const targetTextNoSpace = targetText.replace(/\s+/g, '');
      const chosungText = getInitialSound(targetTextNoSpace);

      const matchQuery = keywords.every(kw => {
        const kwNoSpace = kw.replace(/\s+/g, '');
        return targetText.includes(kw) || targetTextNoSpace.includes(kwNoSpace) || chosungText.includes(kwNoSpace);
      });

      return matchCat && matchQuery;
    });
  }

  function filterExcelCatalog(registeredItems, queryStr, forceShowAll) {
    const catalog = getExcelCatalog();
    if (!catalog || catalog.length === 0) return { matches: [], totalMatches: 0 };

    const rawQ = (queryStr || '').toLowerCase().trim();
    if (!rawQ && !forceShowAll) return { matches: [], totalMatches: 0 };

    const regNames = new Set((registeredItems || []).map(i => (i.name || '').trim().toLowerCase()));
    const keywords = rawQ.split(/\s+/).filter(Boolean);

    const matched = [];
    let totalMatches = 0;

    for (let i = 0; i < catalog.length; i++) {
      const entry = catalog[i];
      const nameStr = (entry.n || '').trim();
      if (!nameStr) continue;
      if (regNames.has(nameStr.toLowerCase())) continue;

      if (rawQ) {
        const lowerName = nameStr.toLowerCase();
        const noSpace = lowerName.replace(/\s+/g, '');
        const chosung = getInitialSound(noSpace);
        const bc = String(entry.b || '').toLowerCase();

        const ok = keywords.every(kw => {
          const kwNoSpace = kw.replace(/\s+/g, '');
          return lowerName.includes(kw) || noSpace.includes(kwNoSpace) || chosung.includes(kwNoSpace) || bc.includes(kwNoSpace);
        });
        if (!ok) continue;
      }

      totalMatches++;
      if (matched.length < excelPageLimit) {
        matched.push(entry);
      }
    }

    return { matches: matched, totalMatches };
  }

  function render(containerId) {
    const container = document.getElementById(containerId || 'module-content');
    if (!container) return;

    const items = getStorageData();
    const zones = getZones();
    const catalogCount = getExcelCatalog().length;

    const filtered = filterRegisteredItems(items, activeCategory, searchQuery);
    const sortedFiltered = sortItems(filtered);
    const excelResult = filterExcelCatalog(items, searchQuery, activeCategory === 'EXCEL_DB');

    const html = `
      <div class="module-header flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 bg-white dark:bg-slate-900 p-4 sm:p-5 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm mb-4 sm:mb-5">
        <div class="space-y-1">
          <h2 class="text-lg sm:text-xl font-extrabold text-slate-900 dark:text-white tracking-tight flex items-center gap-2 flex-wrap">
            <i class="fas fa-boxes-packing text-blue-600"></i>
            <span>💊 일반약 위치 관리 & 엑셀 품목 검색</span>
            <span class="text-xs font-extrabold px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200">초기 엑셀 ${catalogCount.toLocaleString()}품목 탑재 완료</span>
          </h2>
          <p class="text-xs text-slate-500 dark:text-slate-400 leading-relaxed break-keep">
            코페이 POS에서 지원하지 않는 <strong>매대 사진(원거리+근거리) · 번호식 매대 구역(1번~12번) · 창고 여분박스 위치</strong>를 실시간 공유합니다.
          </p>
        </div>
        <div class="flex flex-wrap items-center gap-2 w-full sm:w-auto">
          <label for="med-excel-upload-input" class="flex-1 sm:flex-initial btn font-bold text-xs px-3.5 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white shadow-sm transition flex items-center justify-center gap-1.5 whitespace-nowrap cursor-pointer" title="코페이 POS 또는 약국 품목 엑셀/CSV 일괄 추가 업로드">
            <i class="fas fa-file-excel"></i>
            <span>📥 엑셀 품목 추가 업로드</span>
          </label>
          <input type="file" id="med-excel-upload-input" accept=".csv,.txt,.xlsx,.xls" style="display:none;" onchange="MedicineLocationModule.handleExcelFileUpload(this)">

          <button type="button" class="flex-1 sm:flex-initial btn btn-primary font-bold text-xs px-3.5 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-500 text-white shadow-md transition flex items-center justify-center gap-1.5 whitespace-nowrap" onclick="MedicineLocationModule.openCreateModal()">
            <i class="fas fa-plus-circle"></i>
            <span>+ 새 약품 위치 등록</span>
          </button>
        </div>
      </div>

      <!-- 🔍 1. 실시간 통합 검색 바 (초성 검색 & 엑셀 DB 9,610품목 연동) -->
      <div class="bg-white dark:bg-slate-900 p-3.5 sm:p-4 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm mb-4 sm:mb-5 space-y-3">
        <div class="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-2.5">
          <div class="relative flex-grow">
            <i class="fas fa-search absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 text-xs"></i>
            <input type="text" id="med-search-input" value="${escapeHTML(searchQuery)}" oninput="MedicineLocationModule.handleSearch(this.value)" placeholder="약품명, 초성(예: ㅌㅇㄹㄴ), 매대 번호(예: 1번, 3번), 바코드 검색... (엑셀 9,610품목 동시 검색)" class="w-full bg-slate-50 dark:bg-slate-950 border border-slate-300 dark:border-slate-700 rounded-xl pl-8 pr-8 py-2.5 text-xs font-bold text-slate-900 dark:text-white outline-none focus:border-blue-500">
            <button id="med-search-clear-btn" onclick="MedicineLocationModule.handleSearch('')" style="display:${searchQuery ? 'block' : 'none'};" class="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-xs">✖</button>
          </div>

          <div class="flex items-center gap-2 flex-wrap">
            <div class="flex items-center gap-1.5 text-xs font-bold text-slate-600 dark:text-slate-300 bg-blue-50 dark:bg-slate-800 px-3 py-1.5 rounded-xl whitespace-nowrap border border-blue-200">
              <i class="fas fa-map-marker-alt text-blue-600"></i>
              <span>위치등록 완료: <strong class="text-blue-600 dark:text-blue-400 text-sm font-black">${items.length}</strong>개</span>
            </div>
            <button type="button" onclick="MedicineLocationModule.filterCategory('EXCEL_DB')" class="flex items-center gap-1.5 text-xs font-bold px-3 py-1.5 rounded-xl whitespace-nowrap border transition ${activeCategory === 'EXCEL_DB' ? 'bg-emerald-600 text-white border-emerald-600 shadow' : 'bg-emerald-50 text-emerald-800 border-emerald-200 hover:bg-emerald-100'}">
              <i class="fas fa-database"></i>
              <span>📋 사전탑재 엑셀품목 (${catalogCount.toLocaleString()}개)</span>
            </button>
          </div>
        </div>

        <!-- 🎯 2. 번호식 구역별 퀵 필터 탭 (1번 구역 ~ 12번 구역 + 카운터/냉장고/창고) -->
        <div class="flex items-center gap-1.5 overflow-x-auto pb-1 no-scrollbar text-xs font-bold whitespace-nowrap pt-2 border-t border-slate-100 dark:border-slate-800">
          <button type="button" onclick="MedicineLocationModule.filterCategory('ALL')" class="px-3 py-1.5 rounded-full text-xs font-black transition ${activeCategory === 'ALL' ? 'bg-blue-600 text-white shadow' : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400 hover:bg-slate-200'}">
            전체 위치등록 (${items.length})
          </button>
          ${zones.map(z => {
            const count = items.filter(i => i.zoneId === z.id).length;
            const isActive = activeCategory === z.id;
            return `
              <button type="button" onclick="MedicineLocationModule.filterCategory('${z.id}')" class="px-3 py-1.5 rounded-full text-xs font-bold transition border ${isActive ? 'bg-blue-600 text-white border-blue-600 shadow' : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300 border-slate-200 dark:border-slate-700 hover:bg-slate-200'}">
                <i class="fas ${z.icon}"></i> ${escapeHTML(z.name)} (${count})
              </button>
            `;
          }).join('')}
          <button type="button" onclick="MedicineLocationModule.addNextNumberZone()" class="px-2.5 py-1.5 rounded-full text-xs font-extrabold bg-slate-50 text-blue-600 border border-dashed border-blue-300 hover:bg-blue-50 transition" title="새 매대 번호 구역 추가">
            + 구역번호 추가
          </button>
        </div>
      </div>

      <!-- 🖼️ 3. 약품 카드 그리드 & 엑셀 품목 매칭 컨테이너 -->
      <div id="med-card-grid-container">
        ${renderCombinedContentHTML(sortedFiltered, items.length, excelResult)}
      </div>

      <!-- 📝 4. 신규 등록 / 위치 변경 모달 -->
      <div id="med-location-modal" class="modal-overlay" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(15,23,42,0.75); backdrop-filter:blur(4px); z-index:99999; justify-content:center; align-items:center; padding:16px;">
        <div class="modal-card shadow-2xl" style="background:#ffffff; border-radius:24px; max-width:560px; width:100%; max-height:92vh; overflow-y:auto; padding:26px; position:relative; box-sizing:border-box;">
          <button type="button" onclick="MedicineLocationModule.closeModal()" style="position:absolute; top:20px; right:20px; background:#f1f5f9; border:none; width:34px; height:34px; border-radius:50%; font-size:16px; color:#64748b; cursor:pointer;"><i class="fas fa-times"></i></button>
          
          <div style="display:flex; align-items:center; gap:10px; border-bottom:1.5px solid #f1f5f9; padding-bottom:14px; margin-bottom:18px;">
            <div style="width:40px; height:40px; border-radius:12px; background:#eff6ff; color:#2563eb; display:flex; align-items:center; justify-content:center; font-size:20px;">
              <i class="fas fa-camera-retro"></i>
            </div>
            <div>
              <h3 id="med-modal-title" style="font-size:18px; font-weight:800; color:#0f172a; margin:0;">일반약 보관 위치 등록</h3>
              <p style="font-size:12px; color:#64748b; margin:2px 0 0 0;">매대 번호(1번~12번)와 현장 사진(원거리+근거리)을 등록하세요.</p>
            </div>
          </div>

          <form onsubmit="MedicineLocationModule.handleSubmit(event)">
            <input type="hidden" id="med-item-id" value="" />
            <input type="hidden" id="med-edit-type" value="TYPO_FIX" />

            <!-- ✏️ 수정 모드 선택 박스 (기존 항목 수정 시에만 표시) -->
            <div id="med-edit-mode-box" class="mb-3" style="display:none; background:#f8fafc; border:1.5px solid #cbd5e1; border-radius:14px; padding:12px;">
              <div style="font-size:12px; font-weight:800; color:#334155; margin-bottom:8px; display:flex; align-items:center; gap:5px;">
                <i class="fas fa-sliders-h text-blue-600"></i> <span>작업 유형 선택</span>
              </div>
              <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px;">
                <button type="button" id="med-mode-typo-btn" onclick="MedicineLocationModule.setEditMode('TYPO_FIX')" style="padding:9px 10px; border-radius:10px; font-size:12.5px; font-weight:800; border:2px solid #2563eb; background:#eff6ff; color:#1d4ed8; cursor:pointer; text-align:center; transition:all 0.15s;">
                  ✏️ 오타·내용 수정<br><span style="font-size:10.5px; font-weight:600; opacity:0.85;">(이력 추가 없이 즉시 수정)</span>
                </button>
                <button type="button" id="med-mode-move-btn" onclick="MedicineLocationModule.setEditMode('LOCATION_MOVE')" style="padding:9px 10px; border-radius:10px; font-size:12.5px; font-weight:800; border:1.5px solid #cbd5e1; background:#ffffff; color:#64748b; cursor:pointer; text-align:center; transition:all 0.15s;">
                  🔄 보관 위치 이동<br><span style="font-size:10.5px; font-weight:600; opacity:0.85;">(이전 위치를 이력에 보관)</span>
                </button>
              </div>
            </div>

            <div class="mb-3" style="position:relative;">
              <label class="form-label font-bold" style="font-size:13.5px; color:#334155; margin-bottom:6px;">
                약품명 및 규격 <span style="color:#ef4444;">*</span>
                <span style="font-size:11px; color:#059669; font-weight:700; margin-left:6px;">✨ 이름/초성 입력 시 엑셀 DB 자동완성</span>
              </label>
              <input type="text" id="med-name" required autocomplete="off" oninput="MedicineLocationModule.handleModalNameAutocomplete(this.value)" placeholder="예: 타이레놀정 500mg, 임팩타민 프리미엄 (초성 ㅌㅇㄹㄴ 검색 가능)" style="width:100%; border:1.5px solid #cbd5e1; border-radius:10px; padding:10px 14px; font-size:14px; font-weight:700; outline:none; box-sizing:border-box;" onfocus="this.style.borderColor='#2563eb'" onblur="setTimeout(() => MedicineLocationModule.hideAutocomplete(), 200)" />
              <div id="med-name-autocomplete-box" style="display:none; position:absolute; left:0; right:0; top:72px; background:#ffffff; border:1.5px solid #2563eb; border-radius:12px; max-height:210px; overflow-y:auto; z-index:100; box-shadow:0 10px 25px rgba(15,23,42,0.18);"></div>
            </div>

            <div class="mb-3">
              <label class="form-label font-bold" style="font-size:13.5px; color:#334155; margin-bottom:6px;">매대 구역 번호 선택 <span style="color:#ef4444;">*</span></label>
              <select id="med-zone-id" required style="width:100%; border:1.5px solid #cbd5e1; border-radius:10px; padding:10px 14px; font-size:14px; font-weight:700; outline:none; background:#ffffff; box-sizing:border-box;">
                ${zones.map(z => `<option value="${z.id}">${escapeHTML(z.name)}</option>`).join('')}
              </select>
            </div>

            <div class="mb-3">
              <label class="form-label font-bold" style="font-size:13.5px; color:#334155; margin-bottom:6px;">상세 매대 위치 / 창고 여분박스 위치 <span style="font-weight:normal; color:#64748b; font-size:12px;">(선택)</span></label>
              <input type="text" id="med-location-detail" placeholder="예: 3번 구역 2단 왼쪽 / 여분 박스는 상단 찬장 3번" style="width:100%; border:1.5px solid #cbd5e1; border-radius:10px; padding:10px 14px; font-size:14px; font-weight:700; outline:none; box-sizing:border-box;" onfocus="this.style.borderColor='#2563eb'" onblur="this.style.borderColor='#cbd5e1'" />
            </div>

            <div class="mb-3">
              <label class="form-label font-bold" style="font-size:13.5px; color:#334155; margin-bottom:6px;">코페이 POS 전산위치 / 참고 판매가·바코드 메모 <span style="font-weight:normal; color:#64748b; font-size:12px;">(선택)</span></label>
              <input type="text" id="med-pos-memo" placeholder="예: 코페이 전산위치 3-2 / 판매가 3,000원 / 전산수량 보정메모" style="width:100%; border:1.5px solid #cbd5e1; border-radius:10px; padding:9px 14px; font-size:13px; font-weight:600; outline:none; box-sizing:border-box; background:#f8fafc;" />
            </div>

            <div class="mb-3">
              <label class="form-label font-bold" style="font-size:13.5px; color:#334155; margin-bottom:6px;">보관 위치 사진 첨부 <span style="font-weight:normal; color:#2563eb; font-size:12px; font-weight:700;">(원거리 매대샷 + 근거리 상세샷 최대 5장)</span></label>
              
              <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px; width:100%;">
                <!-- 📸 1. 바로 카메라 촬영 버튼 (스마트폰 직통: capture="environment") -->
                <label for="med-photo-camera" style="display:flex; flex-direction:column; align-items:center; justify-content:center; gap:6px; border-radius:14px; border:2px dashed #2563eb; padding:16px 10px; background:#eff6ff; color:#1d4ed8; font-size:13px; font-weight:800; cursor:pointer; text-align:center;">
                  <i class="fas fa-camera" style="font-size:22px; color:#2563eb;"></i>
                  <span>📸 바로 카메라 촬영</span>
                  <span style="font-size:10px; color:#3b82f6;">(모바일 현장 촬영)</span>
                </label>
                <input type="file" id="med-photo-camera" accept="image/*" capture="environment" style="display:none;" onchange="MedicineLocationModule.handlePhotoSelect(this)">

                <!-- 📁 2. 앨범/사진 선택 버튼 -->
                <label for="med-photo-gallery" style="display:flex; flex-direction:column; align-items:center; justify-content:center; gap:6px; border-radius:14px; border:2px dashed #94a3b8; padding:16px 10px; background:#f8fafc; color:#475569; font-size:13px; font-weight:800; cursor:pointer; text-align:center;">
                  <i class="fas fa-images" style="font-size:22px; color:#64748b;"></i>
                  <span>📁 앨범/사진 선택</span>
                  <span style="font-size:10px; color:#94a3b8;">(갤러리/PC 다중선택)</span>
                </label>
                <input type="file" id="med-photo-gallery" accept="image/*" multiple style="display:none;" onchange="MedicineLocationModule.handlePhotoSelect(this)">
              </div>

              <!-- 📷 다중 사진 미리보기 갤러리 컨테이너 -->
              <div id="med-photo-preview-container" style="display:none; margin-top:12px; background:#f8fafc; padding:12px; border-radius:14px; border:1.5px solid #cbd5e1;">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                  <span id="med-photo-count-text" style="font-size:12px; font-weight:800; color:#1d4ed8;">📷 첨부된 사진 (0/5장)</span>
                  <button type="button" onclick="MedicineLocationModule.resetPhoto()" style="background:#fee2e2; border:1px solid #fca5a5; color:#dc2626; font-size:11px; font-weight:800; padding:3px 8px; border-radius:6px; cursor:pointer;">전체 취소 ✖</button>
                </div>
                <div id="med-photo-thumbnails-grid" style="display:flex; gap:10px; overflow-x:auto; padding-bottom:4px;" class="no-scrollbar"></div>
              </div>
            </div>

            <div class="mb-4">
              <label class="form-label font-bold" style="font-size:13.5px; color:#334155; margin-bottom:6px;">특이사항 및 인수인계 메모 (선택)</label>
              <textarea id="med-notes" rows="2" placeholder="예: 시즌 행사매대 동시 진열 중, 창고 2층에 여분 5박스 보관" style="width:100%; border:1.5px solid #cbd5e1; border-radius:10px; padding:10px 14px; font-size:13.5px; outline:none; box-sizing:border-box;"></textarea>
            </div>

            <div style="display:flex; justify-content:end; gap:8px;">
              <button type="button" onclick="MedicineLocationModule.closeModal()" style="padding:10px 20px; border-radius:10px; background:#f1f5f9; color:#475569; border:none; font-weight:700; cursor:pointer;">취소</button>
              <button type="submit" id="med-submit-btn" style="padding:10px 22px; border-radius:10px; background:linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%); color:#ffffff; border:none; font-weight:800; cursor:pointer; box-shadow:0 4px 12px rgba(37,99,235,0.25);">
                <i class="fas fa-save me-1"></i> 저장하기
              </button>
            </div>
          </form>
        </div>
      </div>

      <!-- 🔍 약품 위치 상세 & 이력 히스토리 팝업 모달 -->
      <div id="med-detail-modal" class="modal-overlay" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(15,23,42,0.8); backdrop-filter:blur(5px); z-index:99999; justify-content:center; align-items:center; padding:16px;">
        <div class="modal-card shadow-2xl" style="background:#ffffff; border-radius:24px; max-width:600px; width:100%; max-height:90vh; overflow-y:auto; padding:28px; position:relative; box-sizing:border-box;">
          <button type="button" onclick="MedicineLocationModule.closeDetailModal()" style="position:absolute; top:20px; right:20px; background:#f1f5f9; border:none; width:34px; height:34px; border-radius:50%; font-size:16px; color:#64748b; cursor:pointer;"><i class="fas fa-times"></i></button>
          <div id="med-detail-modal-content"></div>
        </div>
      </div>
    `;

    container.innerHTML = html;
  }

  function renderMedicineCard(item) {
    const zones = getZones();
    const zone = zones.find(z => z.id === item.zoneId) || { name: item.zoneName || '1번 구역', color: '#2563eb', bg: '#eff6ff', border: '#bfdbfe', icon: 'fa-hashtag' };
    const historyList = (item.history && Array.isArray(item.history)) ? item.history : [];
    const historyCount = historyList.length + 1;
    const allPhotos = (item.photos && Array.isArray(item.photos) && item.photos.length > 0)
      ? item.photos
      : (item.photoUrl ? [item.photoUrl] : []);
    const mainPhoto = allPhotos[0] || '';

    return `
      <div class="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl overflow-hidden flex flex-col justify-between shadow-sm hover:shadow-md transition relative">
        <div>
          <!-- 사진 보관 뷰 영역 -->
          <div class="relative w-full h-40 sm:h-44 bg-slate-100 dark:bg-slate-950 overflow-hidden border-b border-slate-100 dark:border-slate-800 cursor-pointer" onclick="MedicineLocationModule.openDetailModal('${item.id}')">
            ${mainPhoto ? `
              <img src="${mainPhoto}" alt="${escapeHTML(item.name)}" class="w-full h-full object-cover" />
            ` : `
              <div class="flex flex-col items-center justify-center h-full text-slate-400">
                <i class="fas fa-camera text-2xl mb-1"></i>
                <span class="text-xs font-bold">등록된 사진 없음 (터치하여 상세/사진 추가)</span>
              </div>
            `}
            <span class="absolute top-2.5 left-2.5 text-[11px] font-extrabold px-2.5 py-0.5 rounded-full shadow-sm" style="background:${zone.bg}; color:${zone.color}; border:1px solid ${zone.border};">
              <i class="fas ${zone.icon}"></i> ${escapeHTML(zone.name)}
            </span>
            ${allPhotos.length > 1 ? `
              <span class="absolute top-2.5 right-2.5 text-[11px] font-black px-2.5 py-1 rounded-full shadow-md bg-black/75 text-white flex items-center gap-1 backdrop-blur-sm">
                <i class="fas fa-images text-blue-400"></i> 📷 ${allPhotos.length}장
              </span>
            ` : ''}
          </div>

          <!-- 카드 정보 본문 -->
          <div class="p-4 space-y-2.5">
            <h4 class="text-base font-extrabold text-slate-900 dark:text-white leading-tight cursor-pointer break-all" onclick="MedicineLocationModule.openDetailModal('${item.id}')">
              ${escapeHTML(item.name)}
            </h4>

            <div class="bg-blue-50 dark:bg-blue-950/40 border border-blue-200 dark:border-blue-800/80 rounded-xl p-2.5">
              <div class="text-[10px] font-black text-blue-600 dark:text-blue-400 uppercase tracking-wider mb-0.5">📍 매대 보관 위치</div>
              <div class="text-sm font-black text-blue-900 dark:text-blue-200 break-all">
                ${escapeHTML(item.locationDetail)}
              </div>
            </div>

            ${item.posLocationMemo ? `
              <div class="text-xs text-emerald-800 bg-emerald-50 border border-emerald-200 rounded-lg p-2 font-bold leading-snug">
                🏷️ ${escapeHTML(item.posLocationMemo)}
              </div>
            ` : ''}

            ${item.notes ? `
              <p class="text-xs text-amber-800 dark:text-amber-300 bg-amber-50 dark:bg-amber-950/40 border border-amber-200 dark:border-amber-800/80 rounded-lg p-2 font-bold leading-snug">
                💡 ${escapeHTML(item.notes)}
              </p>
            ` : ''}
          </div>
        </div>

        <!-- 카드 하단 관리 메타 바 -->
        <div class="flex items-center justify-between p-3 bg-slate-50 dark:bg-slate-950 border-t border-slate-100 dark:border-slate-800 text-xs">
          <span class="text-[11px] font-medium text-slate-400">
            🕒 <b>${escapeHTML(item.updatedBy || '약국')}</b> · ${escapeHTML(formatLocationDate(item.updatedAt, item.displayDate))}
          </span>

          <div class="flex items-center gap-1.5 flex-wrap">
            <button type="button" onclick="MedicineLocationModule.openEditModal('${item.id}')" class="px-2.5 py-1 rounded-lg bg-slate-100 dark:bg-slate-800 border border-slate-300 dark:border-slate-700 text-slate-700 dark:text-slate-300 font-bold text-xs transition hover:bg-blue-50 hover:border-blue-300">
              <i class="fas fa-edit text-blue-600"></i> 수정/위치이동
            </button>
            <button type="button" onclick="MedicineLocationModule.openDetailModal('${item.id}')" class="px-2.5 py-1 rounded-lg bg-blue-100 dark:bg-blue-950 border border-blue-300 dark:border-blue-700 text-blue-800 dark:text-blue-300 font-bold text-xs transition">
              이력 (${historyCount})
            </button>
            ${window.SheetsSync && window.SheetsSync.getCurrentUser() && window.SheetsSync.getCurrentUser().role === '약국장' ? `
              <button type="button" onclick="MedicineLocationModule.deleteMedicine('${item.id}', '${escapeHTML(item.name).replace(/'/g, "\\'")}')" class="px-2 py-1 rounded-lg bg-rose-100 dark:bg-rose-950 border border-rose-300 dark:border-rose-700 text-rose-700 dark:text-rose-300 font-bold text-xs transition" title="약국장 전용 삭제">
                <i class="fas fa-trash-alt"></i>
              </button>
            ` : ''}
          </div>
        </div>
      </div>
    `;
  }

  function renderExcelCatalogSectionHTML(excelResult) {
    if (!excelResult || !excelResult.matches || excelResult.matches.length === 0) {
      if (activeCategory === 'EXCEL_DB') {
        return `
          <div style="text-align:center; padding:50px 20px; background:#ffffff; border-radius:20px; border:1.5px dashed #cbd5e1;">
            <p style="font-size:14px; font-weight:800; color:#475569;">검색어와 일치하는 엑셀 사전등록 품목이 없습니다.</p>
          </div>
        `;
      }
      return '';
    }

    return `
      <div class="mt-5 bg-white dark:bg-slate-900 rounded-2xl border border-emerald-200 shadow-sm p-4">
        <div class="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 pb-3 mb-3 border-b border-slate-100">
          <div>
            <h3 class="text-sm sm:text-base font-extrabold text-slate-900 flex items-center gap-2">
              <span class="px-2.5 py-0.5 rounded-lg bg-emerald-100 text-emerald-800 text-xs font-black">📋 초기 엑셀 품목 DB</span>
              <span>원클릭 위치·사진 등록 대기 품목 (검색결과 ${excelResult.totalMatches.toLocaleString()}건)</span>
            </h3>
            <p class="text-xs text-slate-500 mt-0.5">품목의 <b>[📸 위치·사진 바로 등록]</b> 버튼을 누르면 약품명을 타자 칠 필요 없이 매대 번호와 사진만 바로 등록됩니다.</p>
          </div>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-2.5">
          ${excelResult.matches.map(entry => {
            const safeName = escapeHTML(entry.n || '');
            const jsName = (entry.n || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
            const priceText = entry.p ? `${Number(entry.p).toLocaleString()}원` : '';
            const barcodeText = entry.b ? `바코드: ${escapeHTML(entry.b)}` : '';
            const memoText = [priceText ? `참고판매가 ${priceText}` : '', barcodeText].filter(Boolean).join(' / ');
            const jsMemo = memoText.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
            return `
              <div class="flex items-center justify-between gap-2 p-3 rounded-xl border border-slate-200 bg-slate-50/70 hover:bg-blue-50/50 hover:border-blue-300 transition">
                <div class="min-w-0 flex-1">
                  <div class="text-xs sm:text-sm font-extrabold text-slate-900 truncate" title="${safeName}">${safeName}</div>
                  <div class="flex items-center gap-1.5 mt-1 flex-wrap">
                    <span class="text-[10px] font-bold px-1.5 py-0.5 rounded bg-slate-200 text-slate-700">${escapeHTML(entry.s || '엑셀DB')}</span>
                    ${priceText ? `<span class="text-[11px] font-extrabold text-emerald-700">${priceText}</span>` : ''}
                    ${entry.b ? `<span class="text-[10px] text-slate-400 font-mono">${escapeHTML(entry.b)}</span>` : ''}
                  </div>
                </div>
                <button type="button" onclick="MedicineLocationModule.openCreateModalWithPreset('${jsName}', '${jsMemo}')" class="px-2.5 py-1.5 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-extrabold text-xs whitespace-nowrap shadow-sm flex items-center gap-1 shrink-0">
                  <i class="fas fa-camera"></i>
                  <span>위치·사진 등록</span>
                </button>
              </div>
            `;
          }).join('')}
        </div>

        ${excelResult.totalMatches > excelResult.matches.length ? `
          <div class="text-center mt-4">
            <button type="button" onclick="MedicineLocationModule.loadMoreExcelItems()" class="px-4 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-extrabold text-xs border border-slate-300 transition">
              + 엑셀 품목 더 보기 (${excelResult.matches.length} / ${excelResult.totalMatches.toLocaleString()}건)
            </button>
          </div>
        ` : ''}
      </div>
    `;
  }

  function renderCombinedContentHTML(sortedFiltered, totalRegisteredCount, excelResult) {
    if (activeCategory === 'EXCEL_DB') {
      return renderExcelCatalogSectionHTML(excelResult);
    }

    let cardsHtml = '';
    if (!sortedFiltered || sortedFiltered.length === 0) {
      cardsHtml = `
        <div style="text-align:center; padding:44px 20px; background:#ffffff; border-radius:20px; border:1.5px dashed #cbd5e1; margin-top:6px;">
          <div style="width:54px; height:54px; border-radius:50%; background:#eff6ff; color:#2563eb; display:flex; align-items:center; justify-content:center; font-size:24px; margin:0 auto 12px auto;">
            <i class="fas fa-box-open"></i>
          </div>
          <h3 style="font-size:16px; font-weight:800; color:#0f172a; margin-bottom:6px;">
            ${totalRegisteredCount === 0 ? '아직 사진/매대위치가 확정된 일반약이 없습니다.' : '일치하는 매대 위치 등록 약품이 없습니다.'}
          </h3>
          <p style="font-size:13px; color:#64748b; margin-bottom:14px;">
            위 검색창에 약 이름이나 초성(예: <b>ㅌㅇㄹㄴ</b>)을 입력하시면 <b>사전 탑재된 9,610개 엑셀 품목</b>이 즉시 나타나 사진만 찍어 3초 만에 등록하실 수 있습니다!
          </p>
          <div style="display:flex; justify-content:center; gap:8px; flex-wrap:wrap;">
            <button type="button" class="btn btn-primary font-bold" onclick="MedicineLocationModule.openCreateModal()" style="border-radius:10px; padding:9px 18px; font-size:13px;">
              <i class="fas fa-plus"></i> 새 약품 위치 등록하기
            </button>
            <button type="button" onclick="MedicineLocationModule.filterCategory('EXCEL_DB')" style="border-radius:10px; padding:9px 18px; font-size:13px; font-weight:800; background:#ecfdf5; color:#047857; border:1.5px solid #6ee7b7; cursor:pointer;">
              <i class="fas fa-list-ul"></i> 📋 초기 엑셀 9,610품목 전체 보기
            </button>
          </div>
        </div>
      `;
    } else {
      cardsHtml = `
        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3.5 sm:gap-4">
          ${sortedFiltered.map(item => renderMedicineCard(item)).join('')}
        </div>
      `;
    }

    return cardsHtml + renderExcelCatalogSectionHTML(excelResult);
  }

  let selectedMedPhotos = []; // [{ id, data, isNew }]

  async function handlePhotoSelect(inputEl) {
    if (!inputEl || !inputEl.files || inputEl.files.length === 0) return;
    const files = Array.from(inputEl.files);

    const remainingSlots = 5 - selectedMedPhotos.length;
    if (remainingSlots <= 0) {
      alert('⚠️ 사진은 최대 5장까지 첨부할 수 있습니다.');
      inputEl.value = '';
      return;
    }

    const filesToProcess = files.slice(0, remainingSlots);

    for (const file of filesToProcess) {
      try {
        const compressedBase64 = await compressPhotoFile(file);
        if (compressedBase64) {
          selectedMedPhotos.push({
            id: 'med_p_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5),
            data: compressedBase64,
            isNew: true
          });
        }
      } catch (err) {
        console.warn("Photo compress fail:", err);
      }
    }

    inputEl.value = '';
    renderMedPhotoPreviews();
  }

  function compressPhotoFile(file) {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const img = new Image();
          img.onload = () => {
            try {
              const canvas = document.createElement('canvas');
              const MAX_WIDTH = 480;
              let width = img.width || MAX_WIDTH;
              let height = img.height || 360;
              if (width > MAX_WIDTH) {
                height = Math.round((height * MAX_WIDTH) / width);
                width = MAX_WIDTH;
              }
              canvas.width = width;
              canvas.height = height;
              const ctx = canvas.getContext('2d');
              if (ctx) ctx.drawImage(img, 0, 0, width, height);
              resolve(canvas.toDataURL('image/jpeg', 0.4));
            } catch (err) {
              resolve(e.target.result);
            }
          };
          img.onerror = () => resolve('');
          img.src = e.target.result;
        } catch (err) {
          resolve('');
        }
      };
      reader.onerror = () => resolve('');
      reader.readAsDataURL(file);
    });
  }

  function renderMedPhotoPreviews() {
    const container = document.getElementById('med-photo-preview-container');
    const grid = document.getElementById('med-photo-thumbnails-grid');
    const textEl = document.getElementById('med-photo-count-text');
    if (!container || !grid) return;

    if (selectedMedPhotos.length === 0) {
      container.style.display = 'none';
      grid.innerHTML = '';
      return;
    }

    container.style.display = 'block';
    if (textEl) textEl.innerText = `📷 첨부된 사진 (${selectedMedPhotos.length}/5장)`;

    grid.innerHTML = selectedMedPhotos.map((photo, idx) => `
      <div style="position:relative; flex-shrink:0; width:72px; height:72px; border-radius:10px; overflow:hidden; border:1.5px solid #cbd5e1; background:#ffffff; box-shadow:0 1px 3px rgba(0,0,0,0.1);">
        <img src="${photo.data}" alt="사진 ${idx + 1}" style="width:100%; height:100%; object-fit:cover;" />
        <button type="button" onclick="MedicineLocationModule.removeMedPhoto(${idx})" style="position:absolute; top:2px; right:2px; background:rgba(220,38,38,0.9); color:#ffffff; border:none; width:20px; height:20px; border-radius:50%; font-size:11px; font-weight:900; cursor:pointer; display:flex; align-items:center; justify-content:center; box-shadow:0 1px 3px rgba(0,0,0,0.3); z-index:2;">✕</button>
        <span style="position:absolute; bottom:2px; left:2px; background:rgba(15,23,42,0.7); color:#fff; font-size:9px; font-weight:800; padding:1px 4px; border-radius:4px;">${idx + 1}</span>
      </div>
    `).join('');
  }

  function removeMedPhoto(idx) {
    if (idx >= 0 && idx < selectedMedPhotos.length) {
      selectedMedPhotos.splice(idx, 1);
      renderMedPhotoPreviews();
    }
  }

  function resetPhoto() {
    selectedMedPhotos = [];
    const camInput = document.getElementById('med-photo-camera');
    const galInput = document.getElementById('med-photo-gallery');
    if (camInput) camInput.value = '';
    if (galInput) galInput.value = '';
    renderMedPhotoPreviews();
  }

  function setEditMode(mode) {
    const typeInput = document.getElementById('med-edit-type');
    if (typeInput) typeInput.value = mode;
    const typoBtn = document.getElementById('med-mode-typo-btn');
    const moveBtn = document.getElementById('med-mode-move-btn');
    if (typoBtn && moveBtn) {
      if (mode === 'LOCATION_MOVE') {
        moveBtn.style.border = '2px solid #2563eb';
        moveBtn.style.background = '#eff6ff';
        moveBtn.style.color = '#1d4ed8';
        typoBtn.style.border = '1.5px solid #cbd5e1';
        typoBtn.style.background = '#ffffff';
        typoBtn.style.color = '#64748b';
      } else {
        typoBtn.style.border = '2px solid #2563eb';
        typoBtn.style.background = '#eff6ff';
        typoBtn.style.color = '#1d4ed8';
        moveBtn.style.border = '1.5px solid #cbd5e1';
        moveBtn.style.background = '#ffffff';
        moveBtn.style.color = '#64748b';
      }
    }
  }

  function openCreateModal() {
    openCreateModalWithPreset('', '');
  }

  function openCreateModalWithPreset(presetName, presetPosMemo) {
    const zones = getZones();
    const defaultZone = (activeCategory !== 'ALL' && activeCategory !== 'EXCEL_DB') ? activeCategory : (zones[0] ? zones[0].id : 'ZONE_1');
    document.getElementById('med-item-id').value = '';
    document.getElementById('med-modal-title').innerText = presetName ? `📸 [${presetName}] 위치·사진 등록` : '일반약 보관 위치 등록';
    const modeBox = document.getElementById('med-edit-mode-box');
    if (modeBox) modeBox.style.display = 'none';
    setEditMode('TYPO_FIX');
    document.getElementById('med-name').value = presetName || '';
    document.getElementById('med-name').readOnly = false;
    document.getElementById('med-zone-id').value = defaultZone;
    document.getElementById('med-location-detail').value = '';
    const posMemoEl = document.getElementById('med-pos-memo');
    if (posMemoEl) posMemoEl.value = presetPosMemo || '';
    document.getElementById('med-notes').value = '';
    hideAutocomplete();
    resetPhoto();
    document.getElementById('med-location-modal').style.display = 'flex';
  }

  function handleModalNameAutocomplete(val) {
    const box = document.getElementById('med-name-autocomplete-box');
    if (!box) return;
    const q = (val || '').trim().toLowerCase();
    if (q.length < 1) {
      box.style.display = 'none';
      return;
    }

    const catalog = getExcelCatalog();
    const qNoSpace = q.replace(/\s+/g, '');
    const matches = [];
    for (let i = 0; i < catalog.length; i++) {
      const entry = catalog[i];
      const nameStr = (entry.n || '').trim();
      if (!nameStr) continue;
      const lower = nameStr.toLowerCase();
      const noSpace = lower.replace(/\s+/g, '');
      const chosung = getInitialSound(noSpace);
      if (lower.includes(q) || noSpace.includes(qNoSpace) || chosung.includes(qNoSpace)) {
        matches.push(entry);
        if (matches.length >= 12) break;
      }
    }

    if (matches.length === 0) {
      box.style.display = 'none';
      return;
    }

    box.innerHTML = matches.map(entry => {
      const safeName = escapeHTML(entry.n);
      const jsName = (entry.n || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
      const priceText = entry.p ? `${Number(entry.p).toLocaleString()}원` : '';
      const memoText = [priceText ? `참고판매가 ${priceText}` : '', entry.b ? `바코드: ${entry.b}` : ''].filter(Boolean).join(' / ');
      const jsMemo = memoText.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
      return `
        <div onmousedown="MedicineLocationModule.selectAutocompleteItem('${jsName}', '${jsMemo}')" style="padding:9px 14px; border-bottom:1px solid #f1f5f9; cursor:pointer; display:flex; justify-content:space-between; align-items:center; gap:8px;" onmouseover="this.style.background='#eff6ff'" onmouseout="this.style.background='#ffffff'">
          <span style="font-size:13px; font-weight:800; color:#0f172a;">${safeName}</span>
          <span style="font-size:11px; color:#059669; font-weight:700; white-space:nowrap;">${priceText || escapeHTML(entry.s || '')}</span>
        </div>
      `;
    }).join('');
    box.style.display = 'block';
  }

  function selectAutocompleteItem(name, memo) {
    const nameEl = document.getElementById('med-name');
    if (nameEl) nameEl.value = name;
    const posMemoEl = document.getElementById('med-pos-memo');
    if (posMemoEl && !posMemoEl.value && memo) posMemoEl.value = memo;
    hideAutocomplete();
  }

  function hideAutocomplete() {
    const box = document.getElementById('med-name-autocomplete-box');
    if (box) box.style.display = 'none';
  }

  function openEditModal(id, initialMode = 'TYPO_FIX') {
    const items = getStorageData();
    const target = items.find(i => i.id === id);
    if (!target) return;

    const zones = getZones();
    document.getElementById('med-item-id').value = target.id;
    document.getElementById('med-modal-title').innerText = `✏️ [${target.name}] 오타 수정 및 위치 변경`;
    const modeBox = document.getElementById('med-edit-mode-box');
    if (modeBox) modeBox.style.display = 'block';
    setEditMode(initialMode || 'TYPO_FIX');
    document.getElementById('med-name').value = target.name || '';
    document.getElementById('med-name').readOnly = false;
    document.getElementById('med-zone-id').value = target.zoneId || (zones[0] ? zones[0].id : 'ZONE_1');
    document.getElementById('med-location-detail').value = target.locationDetail || '';
    const posMemoEl = document.getElementById('med-pos-memo');
    if (posMemoEl) posMemoEl.value = target.posLocationMemo || '';
    document.getElementById('med-notes').value = target.notes || '';
    resetPhoto();

    const existingPhotos = (target.photos && Array.isArray(target.photos) && target.photos.length > 0) 
      ? target.photos 
      : (target.photoUrl ? [target.photoUrl] : []);

    existingPhotos.forEach((url, idx) => {
      selectedMedPhotos.push({
        id: 'med_exist_' + idx,
        data: url,
        isNew: false
      });
    });
    renderMedPhotoPreviews();

    document.getElementById('med-location-modal').style.display = 'flex';
  }

  function closeModal() {
    document.getElementById('med-location-modal').style.display = 'none';
    hideAutocomplete();
    resetPhoto();
  }

  async function handleSubmit(e) {
    if (e && e.preventDefault) e.preventDefault();
    const currUser = window.SheetsSync ? window.SheetsSync.getCurrentUser() : null;
    if (!currUser) {
      alert("로그인이 필요합니다.");
      return;
    }

    const btn = document.getElementById('med-submit-btn');
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = '<i class="fas fa-spinner fa-spin me-1"></i> 저장 중...';
    }

    try {
      const itemId = document.getElementById('med-item-id') ? document.getElementById('med-item-id').value : '';
      const nameElem = document.getElementById('med-name');
      const name = nameElem ? nameElem.value.trim() : '';
      if (!name) {
        alert('⚠️ 약품명을 입력해 주세요.');
        return;
      }

      const zones = getZones();
      const zoneIdElem = document.getElementById('med-zone-id');
      const zoneId = zoneIdElem ? zoneIdElem.value : 'ZONE_1';
      const zoneObj = zones.find(z => z.id === zoneId);
      const zoneName = zoneObj ? zoneObj.name : '1번 구역';

      const locDetailElem = document.getElementById('med-location-detail');
      const locationDetail = (locDetailElem && locDetailElem.value.trim()) ? locDetailElem.value.trim() : `${zoneName} (보관위치 사진 참조)`;
      const posMemoElem = document.getElementById('med-pos-memo');
      const posLocationMemo = posMemoElem ? posMemoElem.value.trim() : '';
      const notesElem = document.getElementById('med-notes');
      const notes = notesElem ? notesElem.value.trim() : '';

      // 🚀 Cloudinary 1순위 다중 사진 병렬 업로드
      let uploadedUrls = [];
      if (selectedMedPhotos && selectedMedPhotos.length > 0) {
        uploadedUrls = await Promise.all(selectedMedPhotos.map(async (p) => {
          if (!p.isNew && p.data && (p.data.startsWith('http') || p.data.startsWith('assets/'))) {
            return p.data;
          }
          if (window.App && typeof window.App.processAndUploadPhoto === 'function') {
            return await window.App.processAndUploadPhoto(p.data);
          }
          return p.data;
        }));
      }
      uploadedUrls = uploadedUrls.filter(Boolean);
      const mainPhotoUrl = uploadedUrls[0] || '';

      const now = new Date();
      const year = now.getFullYear();
      const month = String(now.getMonth() + 1).padStart(2, '0');
      const date = String(now.getDate()).padStart(2, '0');
      const hours = String(now.getHours()).padStart(2, '0');
      const minutes = String(now.getMinutes()).padStart(2, '0');
      const nowStr = `${year}-${month}-${date} ${hours}:${minutes}`;
      const nowMs = Date.now();
      const items = getStorageData() || [];

      if (itemId) {
        const target = items.find(i => i.id === itemId);
        if (target) {
          const editType = document.getElementById('med-edit-type') ? document.getElementById('med-edit-type').value : 'TYPO_FIX';

          if (editType === 'LOCATION_MOVE') {
            const historyEntry = {
              zoneName: target.zoneName,
              locationDetail: target.locationDetail,
              posLocationMemo: target.posLocationMemo || '',
              photoUrl: target.photoUrl,
              photos: target.photos || (target.photoUrl ? [target.photoUrl] : []),
              updatedBy: target.updatedBy,
              updatedAt: target.updatedAt,
              notes: target.notes
            };
            if (!target.history) target.history = [];
            target.history.unshift(historyEntry);
            if (target.history.length > 5) {
              target.history = target.history.slice(0, 5);
            }
          }

          target.name = name;
          target.zoneId = zoneId;
          target.zoneName = zoneName;
          target.locationDetail = locationDetail;
          target.posLocationMemo = posLocationMemo;
          target.photoUrl = mainPhotoUrl;
          target.photos = uploadedUrls;
          target.notes = notes;
          target.updatedBy = currUser.name;
          target.updatedAt = nowMs;
          target.displayDate = nowStr;
        }
      } else {
        const newItem = {
          id: 'med_' + nowMs,
          name,
          zoneId,
          zoneName,
          locationDetail,
          posLocationMemo,
          photoUrl: mainPhotoUrl,
          photos: uploadedUrls,
          notes,
          updatedBy: currUser.name,
          updatedAt: nowMs,
          displayDate: nowStr,
          history: []
        };
        items.unshift(newItem);
      }

      saveStorageData(items);
      closeModal();
      resetPhoto();
      if (activeCategory === 'EXCEL_DB') activeCategory = 'ALL';
      alert('✅ 성공적으로 저장되었습니다!');
      render('module-content');

    } catch (err) {
      console.error("MedicineLocationModule handleSubmit error:", err);
      alert("⚠️ 저장 처리 중 오류가 발생했습니다: " + err.message);
    } finally {
      const liveBtn = document.getElementById('med-submit-btn');
      if (liveBtn) {
        liveBtn.disabled = false;
        liveBtn.innerHTML = '<i class="fas fa-save me-1"></i> 저장하기';
      }
    }
  }

  // 📥 코페이 / 엑셀 / CSV 파일 일괄 추가 업로드 핸들러
  function handleExcelFileUpload(inputEl) {
    if (!inputEl || !inputEl.files || !inputEl.files[0]) return;
    const file = inputEl.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const text = e.target.result || '';
        const lines = text.split(/\r?\n/).filter(l => l.trim());
        const added = [];
        const existingCatalog = getExcelCatalog();
        const seen = new Set(existingCatalog.map(x => (x.n || '').trim().toLowerCase()));

        lines.forEach((line, idx) => {
          if (idx === 0 && (line.includes('품목명') || line.includes('약품명') || line.includes('상품명'))) return;
          const cols = line.split(/,|\t/).map(c => c.replace(/^"|"$/g, '').trim());
          // 코페이/일반 양식 자동 감지 (품목명 컬럼 탐색)
          let name = '';
          let barcode = '';
          let price = 0;
          if (cols.length >= 3 && /^\d+$/.test(cols[1])) {
            barcode = cols[1];
            name = cols[2];
            price = parseInt(cols[7] || cols[3] || '0', 10) || 0;
          } else {
            name = cols[0] || cols[1] || '';
            price = parseInt(cols[1] || cols[2] || '0', 10) || 0;
          }
          if (name && !seen.has(name.toLowerCase())) {
            seen.add(name.toLowerCase());
            added.push({ n: name, b: barcode, p: price, s: '추가업로드' });
          }
        });

        if (added.length > 0) {
          let prevCustom = [];
          try {
            const raw = localStorage.getItem('mds_custom_excel_otc_v1');
            if (raw) prevCustom = JSON.parse(raw) || [];
          } catch (err) {}
          const merged = [...added, ...prevCustom];
          localStorage.setItem('mds_custom_excel_otc_v1', JSON.stringify(merged));
          alert(`🎉 총 ${added.length.toLocaleString()}개의 신규 품목이 앱 품목 DB에 일괄 추가되었습니다!`);
          activeCategory = 'EXCEL_DB';
          render('module-content');
        } else {
          alert('ℹ️ 이미 초기 엑셀 DB(9,610개)에 모두 등록되어 있는 품목이거나, 텍스트/CSV 형식이 아닙니다.');
        }
      } catch (err) {
        alert('⚠️ 파일 읽기 중 오류가 발생했습니다: ' + err.message);
      } finally {
        inputEl.value = '';
      }
    };
    reader.readAsText(file, 'UTF-8');
  }

  function loadMoreExcelItems() {
    excelPageLimit += 90;
    render('module-content');
  }

  function openDetailModal(id) {
    const items = getStorageData();
    const target = items.find(i => i.id === id);
    if (!target) return;

    const zones = getZones();
    const history = (target.history && Array.isArray(target.history)) ? target.history : [];
    const zone = zones.find(z => z.id === target.zoneId) || { name: target.zoneName || '1번 구역', color: '#2563eb', bg: '#eff6ff', border: '#bfdbfe' };
    const allPhotos = (target.photos && Array.isArray(target.photos) && target.photos.length > 0)
      ? target.photos
      : (target.photoUrl ? [target.photoUrl] : []);

    const contentHtml = `
      <div style="display:flex; align-items:center; justify-content:space-between; border-bottom:1.5px solid #f1f5f9; padding-bottom:14px; margin-bottom:16px; padding-right:50px; flex-wrap:wrap; gap:8px;">
        <div style="display:flex; align-items:center; gap:10px;">
          <div style="width:40px; height:40px; border-radius:12px; background:${zone.bg}; color:${zone.color}; display:flex; align-items:center; justify-content:center; font-size:20px;">
            <i class="fas fa-boxes-stacked"></i>
          </div>
          <div>
            <h3 style="font-size:18px; font-weight:800; color:#0f172a; margin:0;">${escapeHTML(target.name)}</h3>
            <span style="font-size:11.5px; font-weight:800; color:${zone.color};">📍 ${escapeHTML(zone.name)}</span>
          </div>
        </div>
        <div style="display:flex; align-items:center; gap:6px;">
          <button type="button" onclick="MedicineLocationModule.closeDetailModal(); MedicineLocationModule.openEditModal('${target.id}', 'TYPO_FIX');" style="background:#eff6ff; border:1px solid #bfdbfe; border-radius:10px; padding:6px 12px; font-size:12px; font-weight:800; color:#1d4ed8; cursor:pointer; flex-shrink:0;">
            <i class="fas fa-edit me-1"></i> 수정/위치이동
          </button>
          ${window.SheetsSync && window.SheetsSync.getCurrentUser() && window.SheetsSync.getCurrentUser().role === '약국장' ? `
            <button type="button" onclick="MedicineLocationModule.deleteMedicine('${target.id}', '${escapeHTML(target.name).replace(/'/g, "\\'")}')" style="background:#fef2f2; border:1px solid #fecdd3; border-radius:10px; padding:6px 12px; font-size:12px; font-weight:800; color:#ef4444; cursor:pointer; flex-shrink:0;">
              <i class="fas fa-trash-can me-1"></i> 삭제
            </button>
          ` : ''}
        </div>
      </div>

      <!-- 대표 위치 사진 및 갤러리 -->
      ${allPhotos.length > 0 ? `
        <div style="margin-bottom:18px; text-align:center; background:#f8fafc; border:1px solid #e2e8f0; border-radius:14px; padding:12px; position:relative;">
          <img src="${allPhotos[0]}" alt="${escapeHTML(target.name)}" onclick="MedicineLocationModule.openPhoto('${target.id}', 0)" style="max-height:280px; width:100%; object-fit:contain; border-radius:10px; cursor:pointer;" />
          
          ${allPhotos.length > 1 ? `
            <div style="display:flex; gap:8px; overflow-x:auto; margin-top:10px; padding:4px 2px; justify-content:center;" class="no-scrollbar">
              ${allPhotos.map((url, pIdx) => `
                <div onclick="MedicineLocationModule.openPhoto('${target.id}', ${pIdx})" style="position:relative; width:56px; height:56px; border-radius:8px; overflow:hidden; border:2px solid ${pIdx === 0 ? '#2563eb' : '#cbd5e1'}; cursor:pointer; flex-shrink:0;">
                  <img src="${url}" style="width:100%; height:100%; object-fit:cover;" />
                  <span style="position:absolute; bottom:1px; right:1px; background:rgba(0,0,0,0.6); color:#fff; font-size:9px; padding:1px 3px; border-radius:3px;">${pIdx+1}</span>
                </div>
              `).join('')}
            </div>
          ` : ''}

          <div style="margin-top:8px;">
            <button type="button" onclick="MedicineLocationModule.openPhoto('${target.id}', 0)" style="display:inline-flex; align-items:center; gap:5px; background:#ffffff; border:1px solid #cbd5e1; padding:6px 14px; border-radius:10px; font-size:12px; font-weight:800; color:#2563eb; cursor:pointer; box-shadow:0 1px 3px rgba(0,0,0,0.05);">
              <i class="fas fa-expand"></i> 사진 ${allPhotos.length > 1 ? `전체 크게보기 (총 ${allPhotos.length}장)` : '원본 크게보기'}
            </button>
          </div>
        </div>
      ` : ''}

      <!-- 현재 보관 정보 -->
      <div style="background:#eff6ff; border:1.5px solid #bfdbfe; border-radius:14px; padding:14px 16px; margin-bottom:20px;">
        <div style="font-size:11.5px; color:#1e40af; font-weight:700; margin-bottom:4px;">🎯 [현재 보관 위치]</div>
        <div style="font-size:15px; font-weight:800; color:#1d4ed8;">${escapeHTML(target.locationDetail)}</div>
        ${target.posLocationMemo ? `<div style="font-size:12.5px; color:#047857; margin-top:6px; font-weight:700;">🏷️ ${escapeHTML(target.posLocationMemo)}</div>` : ''}
        ${target.notes ? `<div style="font-size:12.5px; color:#475569; margin-top:6px; font-weight:600;">💡 ${escapeHTML(target.notes)}</div>` : ''}
        <div style="font-size:11px; color:#64748b; margin-top:8px; padding-top:6px; border-top:1px solid #dbeafe;">
          작성자: <b>${escapeHTML(target.updatedBy || '약국')}</b> · 작성일시: ${formatLocationDate(target.updatedAt, target.displayDate)}
        </div>
      </div>

      <!-- 위치 이동 히스토리 타임라인 -->
      <h4 style="font-size:14px; font-weight:800; color:#0f172a; margin:0 0 12px 0; display:flex; align-items:center; gap:6px;">
        <i class="fas fa-history text-primary"></i> 위치 변경 이력 추적 (${history.length + 1}건)
      </h4>

      <div style="display:flex; flex-direction:column; gap:10px;">
        <div style="border-left:3px solid #2563eb; padding-left:12px; margin-left:4px;">
          <div style="font-size:11px; color:#2563eb; font-weight:800;">[현재 최신 위치] ${formatLocationDate(target.updatedAt, target.displayDate)}</div>
          <div style="font-size:13px; font-weight:800; color:#0f172a;">${escapeHTML(target.zoneName)} - ${escapeHTML(target.locationDetail)}</div>
          <div style="font-size:11px; color:#64748b;">등록자: ${escapeHTML(target.updatedBy)}</div>
        </div>

        ${history.map((h, idx) => `
          <div style="border-left:3px solid #cbd5e1; padding-left:12px; margin-left:4px; opacity:0.85;">
            <div style="font-size:11px; color:#64748b; font-weight:700;">[과거 이력 ${history.length - idx}] ${formatLocationDate(h.updatedAt, h.displayDate)}</div>
            <div style="font-size:12.5px; font-weight:700; color:#475569;">${escapeHTML(h.zoneName)} - ${escapeHTML(h.locationDetail)}</div>
            <div style="font-size:11px; color:#94a3b8;">작성자: ${escapeHTML(h.updatedBy)}</div>
          </div>
        `).join('')}
      </div>
    `;

    document.getElementById('med-detail-modal-content').innerHTML = contentHtml;
    document.getElementById('med-detail-modal').style.display = 'flex';
  }

  function closeDetailModal() {
    document.getElementById('med-detail-modal').style.display = 'none';
  }

  function handleSearch(val) {
    searchQuery = val;
    excelPageLimit = 60;

    const clearBtn = document.getElementById('med-search-clear-btn');
    if (clearBtn) clearBtn.style.display = val ? 'block' : 'none';

    const gridContainer = document.getElementById('med-card-grid-container');
    if (gridContainer) {
      const items = getStorageData();
      const filtered = filterRegisteredItems(items, activeCategory, searchQuery);
      const excelResult = filterExcelCatalog(items, searchQuery, activeCategory === 'EXCEL_DB');
      gridContainer.innerHTML = renderCombinedContentHTML(sortItems(filtered), items.length, excelResult);
    } else {
      render('module-content');
    }
  }

  function filterCategory(catId) {
    activeCategory = catId;
    excelPageLimit = 60;
    render('module-content');
  }

  function deleteMedicine(id, name) {
    const currUser = window.SheetsSync ? window.SheetsSync.getCurrentUser() : null;
    if (!currUser || currUser.role !== '약국장') {
      alert('🔒 약국장 계정만 약품 위치 데이터를 삭제할 수 있습니다.');
      return;
    }

    if (!confirm(`🗑️ [${name}] 약품 위치 등록을 정말로 삭제하시겠습니까?\n(삭제 후 파이어베이스 클라우드에서도 함께 지워집니다)`)) {
      return;
    }

    const items = getStorageData().filter(i => i.id !== id);
    if (window.SheetsSync && typeof window.SheetsSync.addDeletedId === 'function') {
      window.SheetsSync.addDeletedId(id);
    }
    saveStorageData(items);
    closeDetailModal();
    render('module-content');
  }

  if (typeof window !== 'undefined') {
    window.addEventListener('ssg_cloud_updated', () => {
      const active = window.App && typeof window.App.getActiveModule === 'function' ? window.App.getActiveModule() : '';
      if (active === 'medicine-location') {
        const anyModalOpen = Array.from(document.querySelectorAll('.modal-overlay')).some(m => {
          const disp = window.getComputedStyle(m).display;
          return disp !== 'none' && disp !== '';
        });
        if (!anyModalOpen) {
          render('module-content');
        }
      }
    });
  }

  function openPhoto(id, initialIdx = 0) {
    try {
      const items = getStorageData() || [];
      const target = items.find(i => String(i.id) === String(id));
      if (!target) return;
      const allPhotos = (target.photos && Array.isArray(target.photos) && target.photos.length > 0)
        ? target.photos
        : (target.photoUrl ? [target.photoUrl] : []);
      if (allPhotos.length > 0 && window.App && typeof window.App.openImageLightbox === 'function') {
        window.App.openImageLightbox(allPhotos, target.name || '약품 위치 사진', initialIdx);
      }
    } catch(e) {
      console.warn('openPhoto error:', e);
    }
  }

  return {
    render, openCreateModal, openCreateModalWithPreset, openEditModal, setEditMode, closeModal, openDetailModal, closeDetailModal,
    handlePhotoSelect, removeMedPhoto, resetPhoto, handleSubmit, handleSearch, filterCategory, deleteMedicine, openPhoto,
    addNextNumberZone, handleExcelFileUpload, loadMoreExcelItems, handleModalNameAutocomplete, selectAutocompleteItem, hideAutocomplete
  };
})();
