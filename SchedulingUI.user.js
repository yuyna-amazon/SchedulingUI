// ==UserScript==
// @name         SchedulingUI
// @namespace    https://github.com/yuyna-amazon/SchedulingUI
// @version      15.2
// @description  Amazon Logistics SchedulingUI
// @author       yuyna
// @icon         https://www.google.com/s2/favicons?sz=64&domain=amazon.com
// @match        https://logistics.amazon.co.jp/internal/scheduling/dsps*
// @updateURL    https://raw.githubusercontent.com/yuyna-amazon/SchedulingUI/main/SchedulingUI.user.js
// @downloadURL  https://raw.githubusercontent.com/yuyna-amazon/SchedulingUI/main/SchedulingUI.user.js
// @grant        none
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

newFunction();

function newFunction() {
    (function () {
        'use strict';

        // === スクリプトバージョン（ヘッダーの @version と同期） ===
        const SCRIPT_VERSION = (typeof GM_info !== 'undefined' && GM_info.script && GM_info.script.version)
            ? GM_info.script.version
            : '15.0';

        // === 同一開始時刻あたりの必須上限（超過でハイライト） ===
        const REQUIRED_LIMIT_PER_TIME = 15;

        // === 状態管理 ===
        let currentSSDData = null;
        let currentTimeDataList = [];
        let currentTotals = { accepted: 0, required: 0, proDPAccepted: 0, proDPRequired: 0 };
        let isCalculating = false;
        let debounceTimer = null;
        let observer = null;
        let cachedTable = null;
        let isFutureRequiredDownloading = false;
        let lastCalculatedTime = '';

        // === プリコンパイル正規表現 ===
        const TIME_REGEX = /(\d+):(\d+)\s*(午前|午後|am|pm)/i;
        const TIME_STRICT_REGEX = /^\d{1,2}:\d{2}\s*(午前|午後|am|pm)$/i;
        const DIGIT_REGEX = /^\d+$/;
        const DECIMAL_REGEX = /^\d+(?:\.\d+)?$/;
        const DURATION_REGEX = /^(\d+(?:\.\d+)?)\s*hours?$/i;
        const DATE_REGEX = /(\d+)-(\d+)月/;
        const REQUIRED_REGEX = /text:\s*required\(\)/;
        const SCHEDULED_REGEX = /text:\s*scheduled\(\)/;
        const CYCLE_TEXT_REGEX = /\b(CYCLE[\s_-]*[A-Z0-9_]+)\b/i;

        // === 定数 ===
        const SSD_DEFAULTS = {
            'SSD_1': 1, 'SSD_C1': 1, 'SSD_1_B': 1, 'SSD_2': 1,
            'SSD_3': 1, 'SSD_C3': 1, 'SSD_3_B': 1, 'SSD_4': 1
        };
        const SSD_ADJUSTMENT_DEFAULTS = {
            'SSD_1': 0, 'SSD_C1': 0, 'SSD_1_B': 0, 'SSD_2': 0,
            'SSD_3': 0, 'SSD_C3': 0, 'SSD_3_B': 0, 'SSD_4': 0
        };
        const SSD_SOFT_ADJUSTMENT_DEFAULTS = {
            'SSD_1': 0, 'SSD_C1': 0, 'SSD_1_B': 0, 'SSD_2': 0,
            'SSD_3': 0, 'SSD_C3': 0, 'SSD_3_B': 0, 'SSD_4': 0
        };
        const C1C3_OVERRIDE_DEFAULTS = { 'SSD_C1': 0, 'SSD_C3': 0, 'SSD_C1_1': 0, 'SSD_C1_1B': 0, 'SSD_C3_3B': 0, 'SSD_C3_4': 0, 'CVP_BUF_1': 0, 'CVP_BUF_1B': 0, 'CVP_BUF_3B': 0, 'CVP_BUF_4': 0 };

        const SSD_TIME_RANGES = [
            { key: 'SSD_1', min: 0, max: 360 },
            { key: 'SSD_C1', min: 360, max: 360 }, // 現行仕様維持（手入力override用）
            { key: 'SSD_1_B', min: 360, max: 540 },
            { key: 'SSD_2', min: 540, max: 720 },
            { key: 'SSD_3', min: 720, max: 960 },
            { key: 'SSD_C3', min: 960, max: 960 }, // 現行仕様維持（手入力override用）
            { key: 'SSD_3_B', min: 960, max: 1140 },
            { key: 'SSD_4', min: 1140, max: Infinity }
        ];

        const SSD_LIST = ['SSD_1', 'SSD_1_B', 'SSD_2', 'SSD_3', 'SSD_3_B', 'SSD_4', 'SSD_C1', 'SSD_C3'];
        const SPR_LIST = ['SSD_1', 'SSD_1_B', 'SSD_2', 'SSD_3', 'SSD_3_B', 'SSD_4'];

        // === 未来1週間必須ダウンロード設定 ===
        // 横 = 日付 / 縦 = SSD
        // 1_B / 3_B も含めて出力
        const FUTURE_REQUIRED_EXPORT_KEYS = ['SSD_1', 'SSD_1_B', 'SSD_2', 'SSD_3', 'SSD_3_B', 'SSD_4'];
        const FUTURE_REQUIRED_DAYS = 7;
        const FUTURE_REQUIRED_START_OFFSET = 0; // 0 = 選択中の日付を含む1週間


        // === ユーティリティ ===
        const isServiceTypeName = (name) => name && name.includes('AmFlex');

        const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

        const parseTime = (timeStr) => {
            const match = TIME_REGEX.exec(timeStr);
            if (!match) return 0;
            let hour = +match[1];
            const minute = +match[2];
            const period = match[3].toLowerCase();
            if ((period === '午後' || period === 'pm') && hour !== 12) hour += 12;
            if ((period === '午前' || period === 'am') && hour === 12) hour = 0;
            return hour * 60 + minute;
        };

        // "2 hours 30 minutes" / "3 hours" / "45 minutes" → 時間（小数）
        const parseDurationToHours = (text) => {
            if (!text) return '';
            const t = String(text).trim();
            const h = /(\d+(?:\.\d+)?)\s*(?:hours?|時間)/i.exec(t);
            const m = /(\d+)\s*(?:minutes?|mins?|分)/i.exec(t);
            if (!h && !m) return '';
            const hours = (h ? parseFloat(h[1]) : 0) + (m ? Number(m[1]) / 60 : 0);
            if (!hours) return '';
            return Math.round(hours * 100) / 100;
        };

        const getSSDGroup = (timeMinutes) => {
            for (let i = 0; i < SSD_TIME_RANGES.length; i++) {
                const r = SSD_TIME_RANGES[i];
                if (timeMinutes >= r.min && timeMinutes < r.max) return r.key;
            }
            return null;
        };

        const getShortServiceType = (name) => {
            if (!name) return '-';
            const parts = name.split('_');
            return parts.length >= 2 ? parts.slice(1).join('_') : name;
        };

        const getBaseAccepted = (sk, d, overrides) => (sk === 'SSD_C1' || sk === 'SSD_C3') ? (overrides[sk] || 0) : d.accepted;

        const applyPct = (baseSoft, pct) => Math.round(baseSoft * (1 + pct / 100));

        const applyPctForSSD = (sk, baseSoft, pct) => {
            if (sk === 'SSD_C1' || sk === 'SSD_C3') return Math.round(baseSoft);
            return applyPct(baseSoft, pct);
        };

        // overrides['SSD_C1'] = 1B入力値 (= SSD_1_B に割り当てる枠数)
        // SSD_1_B の Soft は override値そのもの、SSD_C1 の Soft は (受諾×SPR) - override値
        const applySubtract = (sk, soft, overrides, subtract) => {
            if (!subtract) return soft;
            if (sk === 'SSD_1_B') return Math.max(0, soft - (overrides['SSD_C1'] || 0));
            if (sk === 'SSD_3_B') return Math.max(0, soft - (overrides['SSD_C3'] || 0));
            return soft;
        };

        const applySoftAdjustment = (soft, adjust) => Math.max(0, soft + (adjust || 0));

        const getCycleMultiplier = (sk, multipliers) => (sk === 'SSD_C1' || sk === 'SSD_C3') ? 1 : (multipliers[sk] || 1);

        // 入力値固定モードで、入力値が「受諾 × SPR」を超えないようにする上限
        const getFixedInputMax = (dataKey, ssdData, multipliers) => {
            const accepted = ssdData && ssdData[dataKey] ? ssdData[dataKey].accepted : 0;
            const spr = multipliers ? (multipliers[dataKey] || 1) : 1;
            return Math.max(0, Math.round(accepted * spr));
        };

        const capFixedInput = (dataKey, value, ssdData, multipliers) => {
            const max = getFixedInputMax(dataKey, ssdData, multipliers);
            const v = Math.max(0, Number(value) || 0);
            return Math.min(v, max);
        };

        const calculateCycleSoft = (sk, baseValue, buffer, multiplier, pct, overrides, subtract, softAdjust, ssdData, allMultipliers) => {
            if (subtract) {
                const fixedMode = get1B3BFixedMode();
                if (sk === 'SSD_1') {
                    if (fixedMode) {
                        // 入力値固定：SSD_C1_1の値をベースにする（上限は SSD_1 の受諾×SPR）
                        const baseRaw = overrides['SSD_C1_1'] !== undefined ? overrides['SSD_C1_1'] : baseValue;
                        const base = capFixedInput('SSD_1', baseRaw, ssdData, allMultipliers);
                        const maxSoft = getFixedInputMax('SSD_1', ssdData, allMultipliers);
                        const baseSoft = base + buffer * multiplier;
                        const soft = applySoftAdjustment(applyPct(baseSoft, pct), softAdjust);
                        return Math.min(maxSoft, soft);
                    } else {
                        // 減算モード：通常計算 - SSD_C1_1入力値
                        const baseSoft = (baseValue + buffer) * multiplier;
                        const soft = applySoftAdjustment(applyPct(baseSoft, pct), softAdjust);
                        return Math.max(0, soft - (overrides['SSD_C1_1'] || 0));
                    }
                }
                if (sk === 'SSD_1_B') {
                    if (fixedMode) {
                        // 入力値固定：SSD_C1_1Bの値をベースにする
                        const baseRaw = overrides['SSD_C1_1B'] !== undefined ? overrides['SSD_C1_1B'] : baseValue;
                        const base = capFixedInput('SSD_1_B', baseRaw, ssdData, allMultipliers);
                        const maxSoft = getFixedInputMax('SSD_1_B', ssdData, allMultipliers);
                        const baseSoft = base + buffer * multiplier;
                        const soft = applySoftAdjustment(applyPct(baseSoft, pct), softAdjust);
                        return Math.min(maxSoft, soft);
                    } else {
                        // 減算モード：通常計算 - SSD_C1_1B入力値
                        const baseSoft = (baseValue + buffer) * multiplier;
                        const soft = applySoftAdjustment(applyPct(baseSoft, pct), softAdjust);
                        return Math.max(0, soft - (overrides['SSD_C1_1B'] || 0));
                    }
                }
                if (sk === 'SSD_3_B') {
                    if (fixedMode) {
                        // 入力値固定：SSD_C3_3Bの値をベースにする
                        const baseRaw = overrides['SSD_C3_3B'] !== undefined ? overrides['SSD_C3_3B'] : baseValue;
                        const base = capFixedInput('SSD_3_B', baseRaw, ssdData, allMultipliers);
                        const maxSoft = getFixedInputMax('SSD_3_B', ssdData, allMultipliers);
                        const baseSoft = base + buffer * multiplier;
                        const soft = applySoftAdjustment(applyPct(baseSoft, pct), softAdjust);
                        return Math.min(maxSoft, soft);
                    } else {
                        // 減算モード：通常計算 - SSD_C3_3B入力値
                        const baseSoft = (baseValue + buffer) * multiplier;
                        const soft = applySoftAdjustment(applyPct(baseSoft, pct), softAdjust);
                        return Math.max(0, soft - (overrides['SSD_C3_3B'] || 0));
                    }
                }
                if (sk === 'SSD_4') {
                    if (fixedMode) {
                        // 入力値固定：SSD_C3_4の値をベースにする
                        const baseRaw = overrides['SSD_C3_4'] !== undefined ? overrides['SSD_C3_4'] : baseValue;
                        const base = capFixedInput('SSD_4', baseRaw, ssdData, allMultipliers);
                        const maxSoft = getFixedInputMax('SSD_4', ssdData, allMultipliers);
                        const baseSoft = base + buffer * multiplier;
                        const soft = applySoftAdjustment(applyPct(baseSoft, pct), softAdjust);
                        return Math.min(maxSoft, soft);
                    } else {
                        // 減算モード：通常計算 - SSD_C3_4入力値
                        const baseSoft = (baseValue + buffer) * multiplier;
                        const soft = applySoftAdjustment(applyPct(baseSoft, pct), softAdjust);
                        return Math.max(0, soft - (overrides['SSD_C3_4'] || 0));
                    }
                }
                if (sk === 'SSD_C1') {
                    if (fixedMode) {
                        // (SSD_1の受諾×SPR_1 - SSD_1入力値) + (SSD_1_Bの受諾×SPR_1B - SSD_1B入力値) - CVP_Buffer×SPR
                        const acc1 = ssdData && ssdData['SSD_1'] ? ssdData['SSD_1'].accepted : 0;
                        const spr1 = allMultipliers ? (allMultipliers['SSD_1'] || 1) : 1;
                        const acc1B = ssdData && ssdData['SSD_1_B'] ? ssdData['SSD_1_B'].accepted : 0;
                        const spr1B = allMultipliers ? (allMultipliers['SSD_1_B'] || 1) : 1;
                        const ov = overrides || {};
                        const part1 = Math.max(0, Math.round(acc1 * spr1) - (ov['SSD_C1_1'] || 0));
                        const part2 = Math.max(0, Math.round(acc1B * spr1B) - (ov['SSD_C1_1B'] || 0));
                        const bufSub = Math.round((ov['CVP_BUF_1'] || 0) * spr1) + Math.round((ov['CVP_BUF_1B'] || 0) * spr1B);
                        return Math.max(0, part1 + part2 + bufSub);
                    } else {
                        // 減算モード：入力値合算をそのまま表示
                        return overrides['SSD_C1'] || 0;
                    }
                }
                if (sk === 'SSD_C3') {
                    if (fixedMode) {
                        // (SSD_3_Bの受諾×SPR_3B - SSD_3B入力値) + (SSD_4の受諾×SPR_4 - SSD_4入力値) - CVP_Buffer×SPR
                        const acc3B = ssdData && ssdData['SSD_3_B'] ? ssdData['SSD_3_B'].accepted : 0;
                        const spr3B = allMultipliers ? (allMultipliers['SSD_3_B'] || 1) : 1;
                        const acc4 = ssdData && ssdData['SSD_4'] ? ssdData['SSD_4'].accepted : 0;
                        const spr4 = allMultipliers ? (allMultipliers['SSD_4'] || 1) : 1;
                        const ov = overrides || {};
                        const part1 = Math.max(0, Math.round(acc3B * spr3B) - (ov['SSD_C3_3B'] || 0));
                        const part2 = Math.max(0, Math.round(acc4 * spr4) - (ov['SSD_C3_4'] || 0));
                        const bufSub = Math.round((ov['CVP_BUF_3B'] || 0) * spr3B) + Math.round((ov['CVP_BUF_4'] || 0) * spr4);
                        return Math.max(0, part1 + part2 + bufSub);
                    } else {
                        // 減算モード：入力値合算をそのまま表示
                        return overrides['SSD_C3'] || 0;
                    }
                }
            }
            const baseSoft = (baseValue + buffer) * multiplier;
            return applySoftAdjustment(applyPctForSSD(sk, baseSoft, pct), softAdjust);
        };

        // === ストレージ ===
        const getStorage = (key, defaultVal) => {
            try {
                const stored = localStorage.getItem(key);
                if (stored === null) return defaultVal;
                if (typeof defaultVal === 'boolean') return stored === 'true';
                if (typeof defaultVal === 'object') return { ...defaultVal, ...JSON.parse(stored) };
                return stored;
            } catch (e) {
                return defaultVal;
            }
        };

        const setStorage = (key, val) => localStorage.setItem(key, typeof val === 'object' ? JSON.stringify(val) : String(val));

        const getSSDMultipliers = () => getStorage('dsp-ssd-multipliers', SSD_DEFAULTS);
        const saveSSDMultipliers = (m) => setStorage('dsp-ssd-multipliers', m);
        const getSSDAdjustments = () => getStorage('dsp-ssd-adjustments', SSD_ADJUSTMENT_DEFAULTS);
        const saveSSDAdjustments = (a) => setStorage('dsp-ssd-adjustments', a);
        const getSSDSoftAdjustments = () => getStorage('dsp-ssd-soft-adjustments', SSD_SOFT_ADJUSTMENT_DEFAULTS);
        const saveSSDSoftAdjustments = (a) => setStorage('dsp-ssd-soft-adjustments', a);
        const getSoftPct = () => Number(getStorage('dsp-soft-pct', '0'));
        const saveSoftPct = (v) => setStorage('dsp-soft-pct', String(v));
        const getHardPct = () => Number(getStorage('dsp-hard-pct', '10'));
        const saveHardPct = (v) => setStorage('dsp-hard-pct', String(v));
        const getC1C3Overrides = () => getStorage('dsp-c1c3-overrides', C1C3_OVERRIDE_DEFAULTS);
        const saveC1C3Overrides = (v) => setStorage('dsp-c1c3-overrides', v);
        const get1B3BFixedMode = () => true;

        // === 日付 ===
        const getDateForFileName = () => {
            const span = document.querySelector('li.selected .dateText');
            if (span) {
                const match = DATE_REGEX.exec(span.textContent);
                if (match) return match[2].padStart(2, '0') + match[1].padStart(2, '0');
            }
            const d = new Date();
            return String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
        };

        const getDateTabs = () => Array.from(document.querySelectorAll('li[data-bind*="goToDate"]'));

        const getSelectedDateIndex = (tabs = getDateTabs()) => tabs.findIndex(li => li.classList.contains('selected'));

        const getSelectedDateText = () => document.querySelector('li.selected .dateText')?.textContent?.trim() || '';

        const formatSelectedDateLabel = () => {
            const li = document.querySelector('li.selected');
            if (!li) return '';
            const raw = li.querySelector('.dateText')?.textContent?.trim() || '';
            const match = DATE_REGEX.exec(raw);
            if (!match) return raw;
            const day = match[1];
            const month = match[2];
            // 曜日は raw テキストの先頭にある（例: "木 4-6月" → "木"）
            const dowMatch = raw.match(/^([月火水木金土日])/);
            const dow = dowMatch ? '（' + dowMatch[1] + '）' : '';
            return month + '月' + day + '日' + dow;
        };

        const parseMonthDayFromText = (text) => {
            const match = DATE_REGEX.exec(text || '');
            if (!match) return null;
            return {
                day: Number(match[1]),
                month: Number(match[2])
            };
        };

        const resolveBaseDateFromSelectedTab = (selectedText) => {
            const md = parseMonthDayFromText(selectedText);
            if (!md) return new Date();

            const today = new Date();
            const candidates = [
                new Date(today.getFullYear() - 1, md.month - 1, md.day),
                new Date(today.getFullYear(), md.month - 1, md.day),
                new Date(today.getFullYear() + 1, md.month - 1, md.day)
            ];

            candidates.sort((a, b) => Math.abs(a - today) - Math.abs(b - today));
            return candidates[0];
        };

        const formatYMD = (date) => {
            const y = date.getFullYear();
            const m = String(date.getMonth() + 1).padStart(2, '0');
            const d = String(date.getDate()).padStart(2, '0');
            return `${y}/${m}/${d}`;
        };

        const formatFileYMD = (date) => {
            const y = date.getFullYear();
            const m = String(date.getMonth() + 1).padStart(2, '0');
            const d = String(date.getDate()).padStart(2, '0');
            return `${y}${m}${d}`;
        };

        const formatExportServiceType = (name) => (name || '-').replace(/_/g, ' ');

        const getTimeParts = (timeMinutes) => ({
            start: Math.floor(timeMinutes / 60),
            minute: String(timeMinutes % 60).padStart(2, '0')
        });

        const isSameMonthDay = (date, text) => {
            const md = parseMonthDayFromText(text);
            if (!md) return false;
            return date.getMonth() + 1 === md.month && date.getDate() === md.day;
        };

        const findDateTabByDate = (date) => {
            const tabs = getDateTabs();
            return tabs.find(tab => {
                const text = tab.querySelector('.dateText')?.textContent?.trim() || '';
                return isSameMonthDay(date, text);
            }) || null;
        };

        const waitFor = async (predicate, timeout = 5000, interval = 100) => {
            const start = Date.now();
            while (Date.now() - start < timeout) {
                try {
                    if (predicate()) return true;
                } catch (e) { }
                await sleep(interval);
            }
            return false;
        };

        // === テーブル検索 ===
        const findTargetTable = () => {
            if (cachedTable && document.contains(cachedTable)) return cachedTable;

            const tables = document.getElementsByTagName('table');
            for (let i = 0; i < tables.length; i++) {
                const ths = tables[i].getElementsByTagName('th');
                for (let j = 0; j < ths.length; j++) {
                    const thText = ths[j].textContent.trim();
                    const thTitle = ths[j].title || '';
                    if (thText.includes('受諾済み') || thTitle.includes('承諾され')) {
                        cachedTable = tables[i];
                        return cachedTable;
                    }
                }
            }
            return null;
        };

        const getTableDataSignature = () => {
            const table = findTargetTable();
            if (!table) return '';
            const tbody = table.querySelector('tbody');
            if (!tbody) return '';
            return `${tbody.rows.length}__${tbody.textContent.slice(0, 1000)}`;
        };

        const forceClick = (el) => {
            if (!el) return;
            el.scrollIntoView?.({ block: 'nearest', inline: 'center' });
            try { el.click(); } catch (e) { }
            try {
                el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
                el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
                el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
            } catch (e) { }
        };

        const clickDateTabAndWait = async (targetInput, retry = 0) => {
            const target = typeof targetInput === 'number'
                ? getDateTabs()[targetInput]
                : targetInput;
            if (!target) throw new Error('指定日付タブが見つかりません');

            const targetText = target.querySelector('.dateText')?.textContent?.trim() || '';
            const beforeTableSig = getTableDataSignature();

            forceClick(target.querySelector('.dateText') || target);

            const selectedChanged = await waitFor(() => getSelectedDateText() === targetText, 8000, 100);
            if (!selectedChanged) {
                if (retry < 2) return clickDateTabAndWait(targetInput, retry + 1);
                throw new Error('日付切替タイムアウト: ' + targetText);
            }

            // テーブル更新待ち
            await waitFor(() => getTableDataSignature() !== beforeTableSig, 5000, 120);
            await sleep(900);

            // 念のため2回再計算
            calculateAndDisplay();
            await sleep(250);
            calculateAndDisplay();
            await sleep(250);

            return targetText;
        };

        const clickDateTabByDateAndWait = async (date, retry = 0) => {
            const target = findDateTabByDate(date);
            if (!target) throw new Error('指定日付タブが見つかりません: ' + formatYMD(date));

            try {
                return await clickDateTabAndWait(target);
            } catch (err) {
                if (retry < 2) return clickDateTabByDateAndWait(date, retry + 1);
                throw err;
            }
        };

        // === 行データ抽出 ===
        const extractRowData = (row) => {
            let timeText = null, requiredValue = 0, acceptedValue = 0, cycleText = '', blockLength = '';
            const cells = row.cells;
            const numericCandidates = [];

            for (let i = 0; i < cells.length; i++) {
                const cell = cells[i];
                const db = cell.dataset.bind || cell.getAttribute('data-bind') || '';
                const text = cell.textContent.trim();

                if (!timeText && TIME_REGEX.test(text)) timeText = text;
                if (!requiredValue && REQUIRED_REGEX.test(db) && !db.includes('total') && DIGIT_REGEX.test(text)) requiredValue = +text;
                if (!acceptedValue && SCHEDULED_REGEX.test(db) && DIGIT_REGEX.test(text)) acceptedValue = +text;
                if (!cycleText) {
                    const cycleMatch = CYCLE_TEXT_REGEX.exec(text);
                    if (cycleMatch) cycleText = cycleMatch[1].replace(/\s+/g, '_').toUpperCase();
                }
                if (!blockLength) {
                    const durationMatch = DURATION_REGEX.exec(text);
                    if (durationMatch) blockLength = durationMatch[1];
                }
                if (DECIMAL_REGEX.test(text) && !REQUIRED_REGEX.test(db) && !SCHEDULED_REGEX.test(db)) {
                    numericCandidates.push(text);
                }
            }

            if (!requiredValue || !acceptedValue) {
                const els = row.querySelectorAll('[data-bind]');
                for (let i = 0; i < els.length; i++) {
                    const db = els[i].getAttribute('data-bind') || '';
                    const text = els[i].textContent.trim();
                    if (!requiredValue && REQUIRED_REGEX.test(db) && !db.includes('total') && DIGIT_REGEX.test(text)) requiredValue = +text;
                    if (!acceptedValue && SCHEDULED_REGEX.test(db) && DIGIT_REGEX.test(text)) acceptedValue = +text;
                }
            }

            if (!timeText) {
                for (let i = 0; i < cells.length; i++) {
                    const text = cells[i].textContent.trim();
                    if (TIME_STRICT_REGEX.test(text)) {
                        timeText = text;
                        break;
                    }
                }
            }

            if (!blockLength && numericCandidates.length > 0) {
                const filteredCandidates = numericCandidates.filter(text => text !== String(requiredValue) && text !== String(acceptedValue));
                blockLength = filteredCandidates.find(text => text.includes('.')) || filteredCandidates[0] || '';
            }

            return { timeText, requiredValue, acceptedValue, cycleText, blockLength };
        };

        // =====================================================
        // === テーブルからデータ読み取り → 全UI再構築
        // =====================================================
        const calculateAndDisplay = () => {
            if (isCalculating) return;
            isCalculating = true;

            try {
                const targetTable = findTargetTable();
                if (!targetTable) {
                    showError('対象テーブルが見つかりません');
                    return;
                }

                const tbody = targetTable.tBodies[0];
                if (!tbody) return;

                const rows = tbody.rows;
                const rowCount = rows.length;
                const timeDataList = [];
                const ssdData = {};
                SSD_LIST.forEach(k => ssdData[k] = { required: 0, accepted: 0 });

                let totalAccepted = 0;
                let totalRequired = 0;
                let proDPAccepted = 0;
                let proDPRequired = 0;

                const serviceTypeIndices = [];
                for (let i = 0; i < rowCount; i++) {
                    const span = rows[i].querySelector('span.expandable');
                    if (span) {
                        const name = span.textContent.trim();
                        if (isServiceTypeName(name)) {
                            serviceTypeIndices.push({
                                index: i,
                                name,
                                isProDP: name.includes('ProDP')
                            });
                        }
                    }
                }

                // 期間（Length）は timeWindowRow に表示されるため、行位置を先に収集して継承する
                const durationIndices = [];
                for (let i = 0; i < rowCount; i++) {
                    const className = rows[i].className || '';
                    if (!className.includes('timeWindowRow')) continue;
                    const span = rows[i].querySelector('span.expandable');
                    if (!span) continue;
                    const hours = parseDurationToHours(span.textContent);
                    if (hours !== '') durationIndices.push({ index: i, hours });
                }

                let currentServiceType = null;
                let serviceTypeIdx = 0;
                let currentDuration = '';
                let durationIdx = 0;

                for (let i = 0; i < rowCount; i++) {
                    while (serviceTypeIdx < serviceTypeIndices.length && serviceTypeIndices[serviceTypeIdx].index <= i) {
                        currentServiceType = serviceTypeIndices[serviceTypeIdx++];
                        currentDuration = '';
                    }
                    while (durationIdx < durationIndices.length && durationIndices[durationIdx].index <= i) {
                        currentDuration = durationIndices[durationIdx++].hours;
                    }

                    const { timeText, requiredValue, acceptedValue, cycleText, blockLength } = extractRowData(rows[i]);
                    if (!timeText) continue;

                    if (currentServiceType?.isProDP) {
                        proDPAccepted += acceptedValue;
                        proDPRequired += requiredValue;
                        continue;
                    }

                    const serviceTypeName = currentServiceType?.name || '-';
                    const timeMinutes = parseTime(timeText);
                    const ssdGroup = getSSDGroup(timeMinutes);

                    timeDataList.push({
                        time: timeText,
                        timeMinutes,
                        serviceType: serviceTypeName,
                        cycle: cycleText || ssdGroup || '',
                        blockLength: currentDuration !== '' ? currentDuration : blockLength,
                        required: requiredValue,
                        accepted: acceptedValue,
                        ssdGroup
                    });

                    if (ssdGroup) {
                        ssdData[ssdGroup].required += requiredValue;
                        ssdData[ssdGroup].accepted += acceptedValue;
                    }

                    totalAccepted += acceptedValue;
                    totalRequired += requiredValue;
                }

                currentSSDData = ssdData;
                currentTimeDataList = timeDataList;
                currentTotals = { accepted: totalAccepted, required: totalRequired, proDPAccepted, proDPRequired };

                const now = new Date();
                lastCalculatedTime = String(now.getHours()).padStart(2, '0') + ':' + String(now.getMinutes()).padStart(2, '0') + ':' + String(now.getSeconds()).padStart(2, '0');

                renderUI();
            } finally {
                isCalculating = false;
            }
        };

        // =====================================================
        // === 設定変更時：Cycleセルだけ in-place 更新
        // =====================================================
        const refreshCycleValues = () => {
            if (!currentSSDData) return;
            if (!document.getElementById('dsp-main-box')) {
                renderUI();
                return;
            }

            const multipliers = getSSDMultipliers();
            const adjustments = getSSDAdjustments();
            const softAdjustments = getSSDSoftAdjustments();
            const overrides = getC1C3Overrides();
            const pct = getSoftPct();
            const hardPct = getHardPct();
            const subtract = true;

            let totalReq = 0, totalAcc = 0, totalSoft = 0, totalHard = 0;

            for (let ci = 0; ci < SSD_LIST.length; ci++) {
                const sk = SSD_LIST[ci];
                const d = currentSSDData[sk];
                const m = getCycleMultiplier(sk, multipliers);
                const adj = adjustments[sk] || 0;
                const softAdj = softAdjustments[sk] || 0;
                const baseAcc = getBaseAccepted(sk, d, overrides);
                const soft = calculateCycleSoft(sk, baseAcc, adj, m, pct, overrides, subtract, softAdj, currentSSDData, multipliers);
                const hard = Math.round(soft * (1 + hardPct / 100));

                const accEl = document.getElementById('cell-acc-' + sk);
                const softEl = document.getElementById('cell-soft-' + sk);
                const hardEl = document.getElementById('cell-hard-' + sk);

                if (accEl) accEl.innerHTML = (sk === 'SSD_C1' || sk === 'SSD_C3') ? '-' : d.accepted + formatAdjustment(adj);
                if (softEl) softEl.textContent = soft;
                if (hardEl) hardEl.textContent = hard;

                totalReq += d.required;
                totalAcc += d.accepted;
                totalSoft += soft;
                totalHard += hard;
            }

            const totalReqEl = document.getElementById('total-req');
            const totalAccEl = document.getElementById('total-acc');
            const totalSoftEl = document.getElementById('total-soft');
            const totalHardEl = document.getElementById('total-hard');

            if (totalReqEl) totalReqEl.textContent = totalReq;
            if (totalAccEl) totalAccEl.textContent = totalAcc;
            if (totalSoftEl) totalSoftEl.textContent = totalSoft;
            if (totalHardEl) totalHardEl.textContent = totalHard;

            const softTotals = calcSoftTotals();
            const softReqEl = document.getElementById('summary-soft-req');
            const softAccEl = document.getElementById('summary-soft-acc');

            if (softReqEl) softReqEl.textContent = softTotals.fromRequired;
            if (softAccEl) softAccEl.textContent = softTotals.fromAccepted;

            const softGapEl = document.getElementById('summary-soft-gap');
            if (softGapEl) {
                const softGap = softTotals.fromAccepted - softTotals.fromRequired;
                softGapEl.textContent = (softGap >= 0 ? '+' : '') + softGap;
                softGapEl.style.color = softGap >= 0 ? '#4CAF50' : '#f44336';
            }

            const pctEl = document.getElementById('pct-label');
            if (pctEl) {
                pctEl.textContent = pct === 0 ? '' : ' ' + (pct > 0 ? '+' : '') + pct + '%';
                pctEl.style.color = pct > 0 ? '#2196F3' : '#f44336';
            }

            const input1B = document.getElementById('c1c3-SSD_C1');
            const input3B = document.getElementById('c1c3-SSD_C3');
            if (input1B) {
                const max1B = getFixedInputMax('SSD_1_B', currentSSDData, multipliers);
                input1B.max = String(max1B);
                input1B.title = 'Max: ' + max1B + '（受諾×SPR）';
            }
            if (input3B) {
                const max3B = getFixedInputMax('SSD_3_B', currentSSDData, multipliers);
                input3B.max = String(max3B);
                input3B.title = 'Max: ' + max3B + '（受諾×SPR）';
            }
        };

        // === debounce ===
        const debouncedCalculate = (delay = 200) => {
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(calculateAndDisplay, delay);
        };

        // === Observer ===
        const startObserver = () => {
            observer?.disconnect();
            observer = new MutationObserver((mutations) => {
                if (isFutureRequiredDownloading) return;

                for (let i = 0; i < mutations.length; i++) {
                    const m = mutations[i];
                    const target = m.target;

                    if (m.type === 'childList') {
                        if (target.tagName === 'TBODY' || target.tagName === 'TABLE' || target.closest?.('table')) {
                            debouncedCalculate(300);
                            return;
                        }
                    }

                    if (m.type === 'attributes' && m.attributeName === 'class' && target.tagName === 'LI') {
                        debouncedCalculate(400);
                        return;
                    }
                }
            });

            observer.observe(document.body, {
                childList: true,
                subtree: true,
                attributes: true,
                attributeFilter: ['class']
            });
        };

        // === Excel ===
        const downloadExcel = () => {
            if (!currentSSDData) {
                alert('データがありません');
                return;
            }

            const multipliers = getSSDMultipliers();
            const adjustments = getSSDAdjustments();
            const softAdjustments = getSSDSoftAdjustments();
            const overrides = getC1C3Overrides();
            const pct = getSoftPct();
            const hardPct = getHardPct();
            const subtract = true;

            const wsData = [['Station', 'Cycle', 'Soft Caps', 'Hard Caps']];

            for (const ssd of SSD_LIST) {
                const d = currentSSDData[ssd];
                const m = getCycleMultiplier(ssd, multipliers);
                const adj = adjustments[ssd] || 0;
                const softAdj = softAdjustments[ssd] || 0;
                const baseAcc = getBaseAccepted(ssd, d, overrides);
                const soft = calculateCycleSoft(ssd, baseAcc, adj, m, pct, overrides, subtract, softAdj, currentSSDData, multipliers);

                wsData.push(['VFK1', ssd, soft, Math.round(soft * (1 + hardPct / 100))]);
            }

            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            ws['!cols'] = [{ wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 12 }];
            XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

            const fileName = 'SSD_caps_' + getDateForFileName() + '.xlsx';
            XLSX.writeFile(wb, fileName);
            showDownloadNotification(fileName);
        };

        const setFutureDownloadButtonState = (isRunning) => {
            const btn = document.getElementById('dl-future-req-btn');
            if (!btn) return;
            btn.disabled = isRunning;
            btn.textContent = isRunning ? '取得中' : '1W';
            btn.style.opacity = isRunning ? '0.7' : '1';
            btn.style.cursor = isRunning ? 'default' : 'pointer';
            const btn1d = document.getElementById('dl-1d-btn');
            if (btn1d) {
                btn1d.disabled = isRunning;
                btn1d.style.opacity = isRunning ? '0.7' : '1';
                btn1d.style.cursor = isRunning ? 'default' : 'pointer';
            }
        };

        const downloadOneDayExcel = () => {
            if (!currentSSDData || !currentTimeDataList.length) {
                alert('データがありません');
                return;
            }

            const selectedText = getSelectedDateText();
            const baseDate = resolveBaseDateFromSelectedTab(selectedText);

            const exportRows = [];
            currentTimeDataList.forEach(td => {
                const timeParts = getTimeParts(td.timeMinutes);
                exportRows.push([
                    formatYMD(baseDate),
                    formatExportServiceType(td.serviceType),
                    td.blockLength === '' || td.blockLength === undefined ? '' : Number(td.blockLength),
                    timeParts.start,
                    timeParts.minute,
                    td.required,
                    td.accepted
                ]);
            });

            if (exportRows.length === 0) {
                alert('データが取得できませんでした');
                return;
            }

            exportRows.sort((a, b) => {
                if (a[3] !== b[3]) return Number(a[3]) - Number(b[3]);
                return Number(a[4]) - Number(b[4]);
            });

            const wsData = [
                ['Target Date', 'Service Type', 'Length', 'Start', 'Time', '必須', '受諾済み'],
                ...exportRows
            ];

            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            ws['!cols'] = [
                { wch: 14 },
                { wch: 28 },
                { wch: 8 },
                { wch: 8 },
                { wch: 8 },
                { wch: 10 },
                { wch: 10 }
            ];
            XLSX.utils.book_append_sheet(wb, ws, '1D_Required_Detail');

            const fileName = 'SSD_required_1D_' + formatFileYMD(baseDate) + '.xlsx';
            XLSX.writeFile(wb, fileName);
            showDownloadNotification(fileName);
        };

        const downloadFutureRequiredExcel = async () => {
            if (isFutureRequiredDownloading) return;
            if (!currentSSDData) {
                alert('データがありません');
                return;
            }

            const initialTabs = getDateTabs();
            const originalIndex = getSelectedDateIndex(initialTabs);

            if (originalIndex < 0) {
                alert('選択中の日付タブが見つかりません');
                return;
            }

            const originalSelectedText = initialTabs[originalIndex]?.querySelector('.dateText')?.textContent?.trim() || '';

            const baseDate = resolveBaseDateFromSelectedTab(originalSelectedText);

                const exportRows = [];

            isFutureRequiredDownloading = true;
            setFutureDownloadButtonState(true);

            try {
                for (let offset = FUTURE_REQUIRED_START_OFFSET; offset < FUTURE_REQUIRED_START_OFFSET + FUTURE_REQUIRED_DAYS; offset++) {
                    const targetDate = new Date(baseDate);
                    targetDate.setDate(baseDate.getDate() + offset);

                    const targetTab = findDateTabByDate(targetDate);
                    if (!targetTab) break;

                    const displayText = await clickDateTabByDateAndWait(targetDate);

                    currentTimeDataList.forEach(td => {
                        const timeParts = getTimeParts(td.timeMinutes);
                        exportRows.push([
                            formatYMD(targetDate),
                            formatExportServiceType(td.serviceType),
                            td.blockLength === '' || td.blockLength === undefined ? '' : Number(td.blockLength),
                            timeParts.start,
                            timeParts.minute,
                            td.required,
                            td.accepted
                        ]);
                    });
                }

                if (exportRows.length === 0) {
                    alert('未来日付のタブが取得できませんでした');
                    return;
                }

                exportRows.sort((a, b) => {
                    if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
                    if (a[3] !== b[3]) return Number(a[3]) - Number(b[3]);
                    return Number(a[4]) - Number(b[4]);
                });

                const wsData = [
                    ['Target Date', 'Service Type', 'Length', 'Start', 'Time', '必須', '受諾済み'],
                    ...exportRows
                ];

                const wb = XLSX.utils.book_new();
                const ws = XLSX.utils.aoa_to_sheet(wsData);

                ws['!cols'] = [
                    { wch: 14 },
                    { wch: 28 },
                    { wch: 8 },
                    { wch: 8 },
                    { wch: 8 },
                    { wch: 10 },
                    { wch: 10 }
                ];

                XLSX.utils.book_append_sheet(wb, ws, '1W_Required_Detail');

                const firstDate = new Date(baseDate);
                firstDate.setDate(baseDate.getDate() + FUTURE_REQUIRED_START_OFFSET);

                const fileName = 'SSD_required_matrix_' + formatFileYMD(firstDate) + '.xlsx';
                XLSX.writeFile(wb, fileName);
                showDownloadNotification(fileName);

            } catch (err) {
                console.error('[DSP Counter] future required export error', err);
                alert('未来1週間ダウンロードでエラー: ' + err.message);
            } finally {
                try {
                    await clickDateTabByDateAndWait(baseDate);
                } catch (restoreErr) {
                    console.warn('[DSP Counter] 元の日付への復帰に失敗', restoreErr);
                }

                isFutureRequiredDownloading = false;
                setFutureDownloadButtonState(false);
            }
        };

        const showDownloadNotification = (fileName) => {
            const n = document.createElement('div');
            n.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);padding:20px 30px;background:#4CAF50;color:white;border-radius:8px;box-shadow:0 4px 20px rgba(0,0,0,0.3);z-index:10001;font-size:16px;text-align:center;';
            n.innerHTML = '<div style="font-size:24px;margin-bottom:10px;">&#10003;</div><div style="font-weight:bold;">' + fileName + '</div>';
            document.body.appendChild(n);
            setTimeout(() => {
                n.style.opacity = '0';
                n.style.transition = 'opacity 0.3s';
                setTimeout(() => n.remove(), 300);
            }, 1500);
        };

        const showError = (message) => {
            document.getElementById('dsp-main-box')?.remove();
            const box = document.createElement('div');
            box.id = 'dsp-main-box';
            box.style.cssText = 'position:fixed;bottom:20px;right:5px;padding:15px 20px;background:#fff;border:2px solid #f44336;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);z-index:9999;font-size:14px;width:400px;';
            box.innerHTML = '<div style="font-weight:bold;color:#f44336;">エラー</div><div style="margin-top:5px;">' + message + '</div>';
            document.body.appendChild(box);
        };

        const formatAdjustment = (adj) => {
            if (adj === 0) return '';
            return '<span style="color:#f44336;font-weight:bold;margin-left:2px;">' + (adj > 0 ? '+' : '') + adj + '</span>';
        };

        const sectionTitle = (text, color) => '<div style="font-weight:bold;color:' + color + ';padding:4px 2px;margin-bottom:6px;font-size:13px;border-bottom:1px solid #eee;">' + text + '</div>';

        // Soft合計を2種類計算
        const calcSoftTotals = () => {
            if (!currentSSDData) return { fromRequired: 0, fromAccepted: 0 };

            const multipliers = getSSDMultipliers();
            const adjustments = getSSDAdjustments();
            const softAdjustments = getSSDSoftAdjustments();
            const overrides = getC1C3Overrides();
            const pct = getSoftPct();
            const subtract = true;

            let fromRequired = 0;
            let fromAccepted = 0;

            for (let i = 0; i < SSD_LIST.length; i++) {
                const sk = SSD_LIST[i];
                const d = currentSSDData[sk];
                const m = getCycleMultiplier(sk, multipliers);
                const adj = adjustments[sk] || 0;
                const softAdj = softAdjustments[sk] || 0;
                const baseAcc = getBaseAccepted(sk, d, overrides);
                // fromRequired: 必須 × SPR（SSD_C1/C3は除外）
                if (sk !== 'SSD_C1' && sk !== 'SSD_C3') {
                    fromRequired += Math.round(d.required * m);
                }
                fromAccepted += calculateCycleSoft(sk, baseAcc, adj, m, pct, overrides, subtract, softAdj, currentSSDData, multipliers);
            }

            return { fromRequired, fromAccepted };
        };

        // =====================================================
        // === 全UI描画
        // =====================================================
        const renderUI = () => {
            document.getElementById('dsp-main-box')?.remove();
            document.getElementById('dsp-spinner-style')?.remove();

            if (!document.getElementById('dsp-spinner-style')) {
                const style = document.createElement('style');
                style.id = 'dsp-spinner-style';
                style.textContent = 'input[id^="mult-"]::-webkit-inner-spin-button,input[id^="mult-"]::-webkit-outer-spin-button,input[id^="cvp-SSD_"]::-webkit-inner-spin-button,input[id^="cvp-SSD_"]::-webkit-outer-spin-button{-webkit-appearance:none;margin:0;}input[id^="mult-"],input[id^="cvp-SSD_"]{-moz-appearance:textfield;}';
                document.head.appendChild(style);
            }

            const multipliers = getSSDMultipliers();
            const adjustments = getSSDAdjustments();
            const softAdjustments = getSSDSoftAdjustments();
            const overrides = getC1C3Overrides();
            const pct = getSoftPct();
            const hardPct = getHardPct();
            const subtract = true;
            const { accepted: totalAccepted, required: totalRequired, proDPAccepted, proDPRequired } = currentTotals;
            let sprInputRows = '';
            for (let si = 0; si < SPR_LIST.length; si++) {
                const sk2 = SPR_LIST[si];
                sprInputRows +=
                    '<div style="display:grid;grid-template-columns:54px 1fr 1fr 1fr;gap:4px;align-items:center;margin:4px 0;">' +
                    '<label style="font-size:11px;color:#555;font-weight:bold;">' + sk2 + '</label>' +
                    '<input type="number" id="mult-' + sk2 + '" value="' + (multipliers[sk2] || 1).toFixed(1) + '" step="0.1" min="0" max="100"' +
                    ' style="width:100%;padding:4px;border:1px solid #ccc;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;" />' +
                    '<input type="number" id="adj-' + sk2 + '" value="' + (adjustments[sk2] || 0) + '" step="1" min="-999" max="999"' +
                    ' style="width:100%;padding:4px;border:2px solid #ffcdd2;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;color:#f44336;font-weight:bold;" />' +
                    '<input type="number" id="soft-adj-' + sk2 + '" value="' + (softAdjustments[sk2] || 0) + '" step="1" min="-999" max="999"' +
                    ' style="width:100%;padding:4px;border:1px solid #90CAF9;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;color:#2196F3;font-weight:bold;" />' +
                    '</div>';
            }

            // ---- 中パネル：Cycle別行 ----
            let ssdRowsHtml = '';
            let totalReq = 0, totalAcc = 0, totalSoft = 0, totalHard = 0;

            for (let ci = 0; ci < SSD_LIST.length; ci++) {
                const sk = SSD_LIST[ci];
                const d = currentSSDData[sk];
                const m = getCycleMultiplier(sk, multipliers);
                const adj = adjustments[sk] || 0;
                const softAdj = softAdjustments[sk] || 0;
                const baseAcc = getBaseAccepted(sk, d, overrides);
                const soft = calculateCycleSoft(sk, baseAcc, adj, m, pct, overrides, subtract, softAdj, currentSSDData, multipliers);
                const hard = Math.round(soft * (1 + hardPct / 100));

                totalReq += d.required;
                totalAcc += d.accepted;
                totalSoft += soft;
                totalHard += hard;

                ssdRowsHtml +=
                    '<div style="display:grid;grid-template-columns:80px 60px 70px 70px 70px;gap:6px;margin:3px 0;padding:8px;background:' + (sk === 'SSD_C1' || sk === 'SSD_C3' ? '#f0f8ff' : '#e3f2fd') + ';border-radius:3px;align-items:center;">' +
                    '<span style="font-weight:bold;font-size:11px;">' + sk + '</span>' +
                    '<span style="color:#FF9800;font-weight:bold;text-align:center;">' + (sk === 'SSD_C1' || sk === 'SSD_C3' ? '-' : d.required) + '</span>' +
                    '<span id="cell-acc-' + sk + '" style="color:#4CAF50;font-weight:bold;text-align:center;">' + (sk === 'SSD_C1' || sk === 'SSD_C3' ? '-' : d.accepted + formatAdjustment(adj)) + '</span>' +
                    '<span id="cell-soft-' + sk + '" style="color:#2196F3;font-weight:bold;text-align:center;">' + soft + '</span>' +
                    '<span id="cell-hard-' + sk + '" style="color:#9C27B0;font-weight:bold;text-align:center;">' + hard + '</span>' +
                    '</div>';
            }

            ssdRowsHtml +=
                '<div style="display:grid;grid-template-columns:80px 60px 70px 70px 70px;gap:6px;margin:4px 0 3px;padding:8px;background:#fffde7;border-radius:3px;align-items:center;border-top:2px solid #fff176;">' +
                '<span style="font-weight:bold;font-size:11px;color:#333;">合計</span>' +
                '<span id="total-req" style="color:#FF9800;font-weight:bold;text-align:center;">' + totalReq + '</span>' +
                '<span id="total-acc" style="color:#4CAF50;font-weight:bold;text-align:center;">' + totalAcc + '</span>' +
                '<span id="total-soft" style="color:#2196F3;font-weight:bold;text-align:center;">' + totalSoft + '</span>' +
                '<span id="total-hard" style="color:#9C27B0;font-weight:bold;text-align:center;">' + totalHard + '</span>' +
                '</div>';

            // ---- 右パネル：開始時刻別 ----
            const lengthValue = function (v) {
                const n = parseFloat(v);
                return isNaN(n) ? -Infinity : n;
            };
            const sortedTimeData = currentTimeDataList.slice().sort(function (a, b) {
                if (a.timeMinutes !== b.timeMinutes) return a.timeMinutes - b.timeMinutes;
                // 同じ開始時刻はLengthの長い順
                const la = lengthValue(a.blockLength);
                const lb = lengthValue(b.blockLength);
                if (la !== lb) return lb - la;
                return a.serviceType.localeCompare(b.serviceType);
            });

            // 同じ開始時刻ごとの必須合計（超過判定用）
            const requiredByTime = {};
            for (let ri = 0; ri < sortedTimeData.length; ri++) {
                const key = sortedTimeData[ri].timeMinutes;
                requiredByTime[key] = (requiredByTime[key] || 0) + (sortedTimeData[ri].required || 0);
            }

            let timeRowsHtml = sortedTimeData.length === 0 ? '<div style="color:#666;padding:10px;">データなし</div>' : '';
            for (let ti = 0; ti < sortedTimeData.length; ti++) {
                const td = sortedTimeData[ti];
                const lengthText = (td.blockLength === '' || td.blockLength === undefined || td.blockLength === null) ? '-' : td.blockLength;
                const timeTotal = requiredByTime[td.timeMinutes] || 0;
                const isOver = timeTotal > REQUIRED_LIMIT_PER_TIME;
                const rowStyle = isOver
                    ? 'background:#ffebee;border-left:3px solid #f44336;'
                    : 'background:#f5f5f5;';
                timeRowsHtml +=
                    '<div style="display:grid;grid-template-columns:60px 96px 46px 42px 42px;gap:6px;margin:3px 0;padding:6px 8px;' + rowStyle + 'border-radius:3px;align-items:center;"' +
                    (isOver ? ' title="' + td.time + ' の必須合計 ' + timeTotal + ' (上限' + REQUIRED_LIMIT_PER_TIME + '超過)"' : '') + '>' +
                    '<span style="font-weight:bold;font-size:11px;white-space:nowrap;">' + td.time + '</span>' +
                    '<span style="font-size:10px;color:#666;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="' + td.serviceType + '">' + getShortServiceType(td.serviceType) + '</span>' +
                    '<span style="color:#795548;font-weight:bold;text-align:center;font-size:11px;">' + lengthText + '</span>' +
                    '<span style="color:#FF9800;font-weight:bold;text-align:center;">' + td.required + '</span>' +
                    '<span style="color:#4CAF50;font-weight:bold;text-align:center;">' + td.accepted + '</span>' +
                    '</div>';
            }

            // ---- サマリー ----
            const diff = totalAccepted - totalRequired;
            const diffColor = diff >= 0 ? '#4CAF50' : '#f44336';
            const proDPHtml = (proDPAccepted || proDPRequired)
                ? '<span style="font-size:11px;color:#9e9e9e;">※ProDP除外: 受諾 ' + proDPAccepted + '</span>'
                : '';

            const softTotals = calcSoftTotals();
            const pctLabelText = pct === 0 ? '' : ' ' + (pct > 0 ? '+' : '') + pct + '%';
            const pctLabelColor = pct > 0 ? '#2196F3' : '#f44336';

            // ---- 左パネル：CVP入力フィールド (4項目) ----
            const storedOv = getC1C3Overrides();
            const valC1_1 = storedOv['SSD_C1_1'] || 0;
            const valC1_1B = storedOv['SSD_C1_1B'] || 0;
            const valC3_3B = storedOv['SSD_C3_3B'] || 0;
            const valC3_4 = storedOv['SSD_C3_4'] || 0;
            const cvpBuf1 = storedOv['CVP_BUF_1'] || 0;
            const cvpBuf1B = storedOv['CVP_BUF_1B'] || 0;
            const cvpBuf3B = storedOv['CVP_BUF_3B'] || 0;
            const cvpBuf4 = storedOv['CVP_BUF_4'] || 0;
            // 合算値をoverridesに反映
            overrides['SSD_C1'] = valC1_1 + valC1_1B;
            overrides['SSD_C3'] = valC3_3B + valC3_4;
            const c1c3InputHtml =
                '<div style="display:grid;grid-template-columns:60px 1fr 1fr;gap:2px 4px;align-items:center;">' +
                '<span></span><span style="text-align:center;font-size:9px;color:#9C27B0;font-weight:bold;">CVP</span><span style="text-align:center;font-size:9px;color:#f44336;font-weight:bold;">Buffer</span>' +
                '<label style="font-size:10px;color:#555;font-weight:bold;text-align:right;">SSD_1</label>' +
                '<input type="number" id="cvp-SSD_C1_1" value="' + valC1_1 + '" step="1" min="0" max="9999"' +
                ' style="width:100%;padding:4px;border:2px solid #CE93D8;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;color:#9C27B0;font-weight:bold;" />' +
                '<input type="number" id="cvp-buf-1" value="' + cvpBuf1 + '" step="1" min="-999" max="999"' +
                ' style="width:100%;padding:4px;border:2px solid #ffcdd2;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;color:#f44336;font-weight:bold;" />' +
                '<label style="font-size:10px;color:#555;font-weight:bold;text-align:right;">SSD_1_B</label>' +
                '<input type="number" id="cvp-SSD_C1_1B" value="' + valC1_1B + '" step="1" min="0" max="9999"' +
                ' style="width:100%;padding:4px;border:2px solid #CE93D8;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;color:#9C27B0;font-weight:bold;" />' +
                '<input type="number" id="cvp-buf-1b" value="' + cvpBuf1B + '" step="1" min="-999" max="999"' +
                ' style="width:100%;padding:4px;border:2px solid #ffcdd2;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;color:#f44336;font-weight:bold;" />' +
                '<label style="font-size:10px;color:#555;font-weight:bold;text-align:right;">SSD_3_B</label>' +
                '<input type="number" id="cvp-SSD_C3_3B" value="' + valC3_3B + '" step="1" min="0" max="9999"' +
                ' style="width:100%;padding:4px;border:2px solid #CE93D8;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;color:#9C27B0;font-weight:bold;" />' +
                '<input type="number" id="cvp-buf-3b" value="' + cvpBuf3B + '" step="1" min="-999" max="999"' +
                ' style="width:100%;padding:4px;border:2px solid #ffcdd2;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;color:#f44336;font-weight:bold;" />' +
                '<label style="font-size:10px;color:#555;font-weight:bold;text-align:right;">SSD_4</label>' +
                '<input type="number" id="cvp-SSD_C3_4" value="' + valC3_4 + '" step="1" min="0" max="9999"' +
                ' style="width:100%;padding:4px;border:2px solid #CE93D8;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;color:#9C27B0;font-weight:bold;" />' +
                '<input type="number" id="cvp-buf-4" value="' + cvpBuf4 + '" step="1" min="-999" max="999"' +
                ' style="width:100%;padding:4px;border:2px solid #ffcdd2;border-radius:4px;text-align:center;font-size:12px;box-sizing:border-box;color:#f44336;font-weight:bold;" />' +
                '</div>';

            // ---- DOM組み立て ----
            const box = document.createElement('div');
            box.id = 'dsp-main-box';
            box.style.cssText = 'position:fixed;bottom:20px;right:5px;display:flex;align-items:flex-start;background:#fff;border:2px solid #4CAF50;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);z-index:9999;font-size:12px;';

            // 左パネル
            const leftPanel = document.createElement('div');
            leftPanel.style.cssText = 'width:270px;min-width:270px;padding:6px 8px;border-right:2px solid #e3f2fd;max-height:600px;overflow-y:auto;overflow-x:hidden;box-sizing:border-box;';
            leftPanel.innerHTML =
                '<div style="display:flex;align-items:baseline;justify-content:center;gap:8px;margin-bottom:4px;padding-bottom:4px;border-bottom:2px solid #e3f2fd;">' +
                '<span style="font-size:18px;font-weight:bold;color:#333;">' + formatSelectedDateLabel() + '</span>' +
                '<span style="font-size:11px;color:#999;">' + lastCalculatedTime + '</span>' +
                '</div>' +
                '<div style="display:grid;grid-template-columns:54px 1fr 1fr 1fr;gap:4px;margin-bottom:6px;padding:0 2px;font-size:12px;color:#999;font-weight:bold;">' +
                '<span></span><span style="text-align:center;">SPR</span><span style="text-align:center;color:#f44336;">Buffer</span><span style="text-align:center;color:#2196F3;">Adjust</span>' +
                '</div>' +
                sprInputRows +
                '<div style="margin-top:8px;display:grid;grid-template-columns:54px 1fr 1fr 1fr;gap:4px;align-items:center;">' +
                '<button id="spr-file-btn" style="grid-column:1/3;width:100%;height:24px;padding:0 2px;background:#9e9e9e;color:white;border:none;border-radius:4px;cursor:pointer;font-size:9px;font-weight:bold;box-sizing:border-box;line-height:24px;">SPR/CVP参照</button>' +
                '<button id="adj-reset-btn" style="width:100%;height:24px;padding:0 2px;background:#ff8a80;color:white;border:none;border-radius:4px;cursor:pointer;font-size:8px;font-weight:bold;box-sizing:border-box;line-height:11px;display:flex;align-items:center;justify-content:center;text-align:center;">Buffer<br>リセット</button>' +
                '<button id="soft-adj-reset-btn" style="width:100%;height:24px;padding:0 2px;background:#64b5f6;color:white;border:none;border-radius:4px;cursor:pointer;font-size:8px;font-weight:bold;box-sizing:border-box;line-height:11px;display:flex;align-items:center;justify-content:center;text-align:center;">Adjust<br>リセット</button>' +
                '</div>' +
                '<input type="file" id="spr-file-input" accept=".xlsx,.xls" style="display:none;" />' +
                '<div style="margin-top:8px;padding-top:4px;border-top:2px solid #e3f2fd;">' +
                c1c3InputHtml +
                '</div>' +
                '<div style="margin-top:8px;padding-top:4px;border-top:2px solid #e3f2fd;">' +
                '<div style="display:flex;align-items:center;gap:12px;justify-content:space-between;">' +
                '<div>' +
                '<div style="font-weight:bold;color:#2196F3;font-size:13px;margin-bottom:6px;">全Soft調整</div>' +
                '<div style="display:flex;align-items:center;gap:6px;">' +
                '<input type="number" id="soft-pct-input" value="' + pct + '" step="1" min="-100" max="200"' +
                ' style="width:80px;padding:5px;border:2px solid #2196F3;border-radius:4px;text-align:center;font-size:14px;font-weight:bold;box-sizing:border-box;" />' +
                '<span style="font-size:14px;font-weight:bold;color:#555;">%</span>' +
                '</div>' +
                '</div>' +
                '<div>' +
                '<div style="font-weight:bold;color:#9C27B0;font-size:13px;margin-bottom:6px;">Hard調整</div>' +
                '<div style="display:flex;align-items:center;gap:6px;">' +
                '<input type="number" id="hard-pct-input" value="' + getHardPct() + '" step="1" min="0" max="500"' +
                ' style="width:60px;padding:5px;border:2px solid #9C27B0;border-radius:4px;text-align:center;font-size:14px;font-weight:bold;box-sizing:border-box;" />' +
                '<span style="font-size:14px;font-weight:bold;color:#9C27B0;">%</span>' +
                '</div>' +
                '</div>' +
                '</div>' +
                '</div>';

            // 中パネル
            const midPanel = document.createElement('div');
            midPanel.style.cssText = 'width:390px;min-width:390px;padding:6px 8px;border-right:2px solid #e3f2fd;max-height:600px;overflow-y:auto;overflow-x:hidden;box-sizing:border-box;';
            midPanel.innerHTML =
                '<div style="background:#e8f5e9;padding:8px 10px;border-radius:5px;margin-bottom:10px;">' +
                '<div style="display:grid;grid-template-columns:1fr auto;align-items:center;margin:4px 0;">' +
                '<span>必須合計: <strong style="color:#F57C00;">' + totalRequired + '</strong></span>' +
                '<span style="font-size:11px;color:#666;">Soft: <strong id="summary-soft-req" style="color:#2196F3;">' + softTotals.fromRequired + '</strong></span>' +
                '</div>' +
                '<div style="display:grid;grid-template-columns:1fr auto;align-items:center;margin:4px 0;">' +
                '<span>受諾済み: <strong style="color:#2e7d32;">' + totalAccepted + '</strong></span>' +
                '<span style="font-size:11px;color:#666;">Soft: <strong id="summary-soft-acc" style="color:#2196F3;">' + softTotals.fromAccepted + '</strong></span>' +
                '</div>' +
                '<div style="display:grid;grid-template-columns:1fr auto;align-items:center;margin:4px 0;padding-top:5px;border-top:1px solid #c8e6c9;">' +
                '<span>Gap: <strong style="color:' + diffColor + ';">' + (diff >= 0 ? '+' : '') + diff + '</strong></span>' +
                '<span style="font-size:11px;color:#666;">Soft Gap: <strong id="summary-soft-gap" style="color:' + (softTotals.fromAccepted - softTotals.fromRequired >= 0 ? '#4CAF50' : '#f44336') + ';">' + (softTotals.fromAccepted - softTotals.fromRequired >= 0 ? '+' : '') + (softTotals.fromAccepted - softTotals.fromRequired) + '</strong></span>' +
                '</div>' +
                (proDPHtml ? '<div style="margin:2px 0;text-align:right;">' + proDPHtml + '</div>' : '') +
                '</div>' +
                '<div style="display:grid;grid-template-columns:80px 60px 70px 70px 70px;gap:6px;margin-bottom:5px;padding:3px 5px;font-weight:bold;color:#666;font-size:10px;">' +
                '<span>Cycle</span><span style="text-align:center;">必須</span>' +
                '<span style="text-align:center;">受諾</span>' +
                '<span style="text-align:center;">Soft<span id="pct-label" style="font-size:9px;color:' + pctLabelColor + ';">' + pctLabelText + '</span></span>' +
                '<span style="text-align:center;">Hard</span>' +
                '</div>' +
                ssdRowsHtml +
                '<div style="padding-top:10px;border-top:1px solid #ddd;display:flex;gap:6px;">' +
                '<button id="dl-btn" style="flex:5;padding:5px;background:#4CAF50;color:white;border:none;border-radius:5px;cursor:pointer;font-size:12px;font-weight:bold;">Excel download</button>' +
                '<button id="dl-1d-btn" style="flex:1;padding:5px;background:#1E88E5;color:white;border:none;border-radius:5px;cursor:pointer;font-size:11px;font-weight:bold;white-space:nowrap;">1D</button>' +
                '<button id="dl-future-req-btn" style="flex:1;padding:5px;background:#1E88E5;color:white;border:none;border-radius:5px;cursor:pointer;font-size:11px;font-weight:bold;white-space:nowrap;">1W</button>' +
                '</div>';

            // 右パネル
            const rightPanel = document.createElement('div');
            rightPanel.style.cssText = 'width:350px;min-width:350px;padding:6px 8px;overflow-y:auto;overflow-x:hidden;box-sizing:border-box;';
            rightPanel.innerHTML =
                sectionTitle('開始時刻別', '#4CAF50') +
                '<div style="display:grid;grid-template-columns:60px 96px 46px 42px 42px;gap:6px;margin-bottom:5px;padding:3px 8px;font-weight:bold;color:#666;font-size:10px;">' +
                '<span>開始時刻</span><span>サービスタイプ</span><span style="text-align:center;">Length</span><span style="text-align:center;">必須</span><span style="text-align:center;">受諾</span>' +
                '</div>' +
                timeRowsHtml;

            box.appendChild(leftPanel);
            box.appendChild(midPanel);
            box.appendChild(rightPanel);

            // 右上に全体表示/非表示ボタン
            const toggleBtn = document.createElement('button');
            toggleBtn.id = 'dsp-toggle-btn';
            toggleBtn.textContent = '×';
            toggleBtn.style.cssText = 'position:absolute;top:4px;right:8px;background:none;border:none;font-size:18px;cursor:pointer;color:#999;font-weight:bold;line-height:1;z-index:1;';
            box.style.position = 'fixed';
            box.appendChild(toggleBtn);

            // ×ボタンの左隣にバージョン表示
            const versionLabel = document.createElement('span');
            versionLabel.id = 'dsp-version-label';
            versionLabel.textContent = 'v' + SCRIPT_VERSION;
            versionLabel.style.cssText = 'position:absolute;top:7px;right:28px;font-size:10px;color:#bbb;font-weight:bold;line-height:1;z-index:1;pointer-events:none;';
            box.appendChild(versionLabel);

            document.body.appendChild(box);

            var isHidden = false;
            var hideUI = function () {
                leftPanel.style.display = 'none';
                midPanel.style.display = 'none';
                rightPanel.style.display = 'none';
                box.style.border = 'none';
                box.style.boxShadow = 'none';
                box.style.background = 'transparent';
                box.style.bottom = 'auto';
                box.style.top = '10px';
                toggleBtn.textContent = 'Cap計算';
                toggleBtn.style.fontSize = '12px';
                toggleBtn.style.color = '#4CAF50';
                toggleBtn.style.padding = '6px 30px';
                toggleBtn.style.background = '#e8f5e9';
                toggleBtn.style.borderRadius = '4px';
                toggleBtn.style.border = '1px solid #4CAF50';
                toggleBtn.style.whiteSpace = 'nowrap';
                versionLabel.style.display = 'none';
                isHidden = true;
            };
            var showUI = function () {
                leftPanel.style.display = '';
                midPanel.style.display = '';
                rightPanel.style.display = '';
                box.style.border = '2px solid #4CAF50';
                box.style.boxShadow = '0 2px 10px rgba(0,0,0,0.1)';
                box.style.background = '#fff';
                box.style.top = 'auto';
                box.style.bottom = '20px';
                toggleBtn.textContent = '×';
                toggleBtn.style.fontSize = '18px';
                toggleBtn.style.color = '#999';
                toggleBtn.style.padding = '0';
                toggleBtn.style.background = 'none';
                toggleBtn.style.borderRadius = '0';
                toggleBtn.style.border = 'none';
                toggleBtn.style.whiteSpace = '';
                versionLabel.style.display = '';
                isHidden = false;
            };
            toggleBtn.addEventListener('click', function () {
                if (isHidden) { showUI(); } else { hideUI(); }
            });

            setTimeout(function () {
                const h = midPanel.offsetHeight;
                if (h > 0) {
                    rightPanel.style.height = h + 'px';
                    rightPanel.style.maxHeight = h + 'px';
                }
            }, 0);

            // ---- イベント登録 ----
            SPR_LIST.forEach(function (ssd) {
                const multEl = document.getElementById('mult-' + ssd);
                const adjEl = document.getElementById('adj-' + ssd);
                const softAdjEl = document.getElementById('soft-adj-' + ssd);

                if (multEl) {
                    multEl.addEventListener('change', function () {
                        const mv = getSSDMultipliers();
                        mv[ssd] = Math.round((+this.value || 1) * 10) / 10;
                        this.value = mv[ssd].toFixed(1);
                        saveSSDMultipliers(mv);
                        refreshCycleValues();
                    });
                }

                if (adjEl) {
                    adjEl.addEventListener('change', function () {
                        const av = getSSDAdjustments();
                        av[ssd] = +this.value || 0;
                        saveSSDAdjustments(av);
                        refreshCycleValues();
                    });
                }

                if (softAdjEl) {
                    softAdjEl.addEventListener('change', function () {
                        const av = getSSDSoftAdjustments();
                        av[ssd] = +this.value || 0;
                        saveSSDSoftAdjustments(av);
                        refreshCycleValues();
                    });
                }
            });

            var processSPRExcelData = function (data, showAlert) {
                try {
                    var arr = new Uint8Array(data);
                    var wb = XLSX.read(arr, { type: 'array' });
                    var ws = wb.Sheets[wb.SheetNames[0]];
                    var json = XLSX.utils.sheet_to_json(ws, { header: 1 });
                    if (!json || json.length < 7) {
                        if (showAlert) alert('Excelの行数が不足しています（7行以上必要）');
                        return;
                    }
                    // キャッシュに保存
                    try { localStorage.setItem('dsp-spr-excel-cache', JSON.stringify(json)); } catch (e2) { }

                    applySPRFromJson(json, showAlert);
                } catch (err) {
                    if (showAlert) alert('Excel読み込みエラー: ' + err.message);
                }
            };

            var applySPRFromJson = function (json, showAlert) {
                if (!json || json.length < 7) return;
                var headerRow = json[0];
                var selectedText = getSelectedDateText();
                var md = parseMonthDayFromText(selectedText);
                if (!md) {
                    if (showAlert) alert('現在の日付が取得できません');
                    return;
                }
                var colIdx = -1;
                for (var c = 0; c < headerRow.length; c++) {
                    var cell = headerRow[c];
                    if (!cell) continue;
                    var cellDate = null;
                    if (typeof cell === 'number') {
                        cellDate = new Date((cell - 25569) * 86400 * 1000);
                    } else if (typeof cell === 'string') {
                        cellDate = new Date(cell);
                    }
                    if (cellDate && !isNaN(cellDate.getTime())) {
                        if (cellDate.getMonth() + 1 === md.month && cellDate.getDate() === md.day) {
                            colIdx = c;
                            break;
                        }
                    }
                }
                if (colIdx < 0) {
                    // 該当日付なし：SPR/CVP全て0にする
                    var mv = getSSDMultipliers();
                    for (var ri = 0; ri < SPR_LIST.length; ri++) {
                        mv[SPR_LIST[ri]] = 0;
                        var eli = document.getElementById('mult-' + SPR_LIST[ri]);
                        if (eli) eli.value = '0.0';
                    }
                    saveSSDMultipliers(mv);
                    var ov0 = getC1C3Overrides();
                    var cvpKeys0 = ['SSD_C1_1', 'SSD_C1_1B', 'SSD_C3_3B', 'SSD_C3_4'];
                    var cvpIds0 = ['cvp-SSD_C1_1', 'cvp-SSD_C1_1B', 'cvp-SSD_C3_3B', 'cvp-SSD_C3_4'];
                    for (var zi = 0; zi < cvpKeys0.length; zi++) {
                        ov0[cvpKeys0[zi]] = 0;
                        var ze = document.getElementById(cvpIds0[zi]);
                        if (ze) ze.value = 0;
                    }
                    ov0['SSD_C1'] = 0;
                    ov0['SSD_C3'] = 0;
                    saveC1C3Overrides(ov0);
                    refreshCycleValues();
                    if (showAlert) alert('該当日付の列が見つかりません（' + md.month + '月' + md.day + '日）。SPR/CVPを0に設定しました。');
                    return;
                }
                var mv = getSSDMultipliers();
                for (var r = 1; r <= 6 && r < json.length; r++) {
                    var val = json[r][colIdx];
                    var ssd = SPR_LIST[r - 1];
                    if (ssd) {
                        mv[ssd] = (val !== undefined && val !== null && val !== '') ? Math.round(Number(val) * 10) / 10 || 0 : 0;
                        var el = document.getElementById('mult-' + ssd);
                        if (el) el.value = mv[ssd].toFixed(1);
                    }
                }
                saveSSDMultipliers(mv);

                // 11〜16行目: CVP値をc1c3 overridesに反映
                // 行11=SSD_1→SSD_C1_1, 行12=SSD_1_B→SSD_C1_1B, 行15=SSD_3_B→SSD_C3_3B, 行16=SSD_4→SSD_C3_4
                var cvpMapping = [
                    { row: 10, key: 'SSD_C1_1', inputId: 'cvp-SSD_C1_1' },
                    { row: 11, key: 'SSD_C1_1B', inputId: 'cvp-SSD_C1_1B' },
                    { row: 14, key: 'SSD_C3_3B', inputId: 'cvp-SSD_C3_3B' },
                    { row: 15, key: 'SSD_C3_4', inputId: 'cvp-SSD_C3_4' }
                ];
                var ov = getC1C3Overrides();
                for (var ci = 0; ci < cvpMapping.length; ci++) {
                    var map = cvpMapping[ci];
                    if (map.row < json.length) {
                        var cvpVal = json[map.row][colIdx];
                        var numVal = (cvpVal !== undefined && cvpVal !== null && cvpVal !== '') ? Math.floor(Number(cvpVal)) || 0 : 0;
                        ov[map.key] = numVal;
                        var cvpEl = document.getElementById(map.inputId);
                        if (cvpEl) cvpEl.value = numVal;
                    }
                }
                ov['SSD_C1'] = (ov['SSD_C1_1'] || 0) + (ov['SSD_C1_1B'] || 0);
                ov['SSD_C3'] = (ov['SSD_C3_3B'] || 0) + (ov['SSD_C3_4'] || 0);
                saveC1C3Overrides(ov);

                refreshCycleValues();
                if (showAlert) alert('SPR/CVPを反映しました。');
            };

            // 日付変更時にキャッシュからSPR自動反映
            var applySPRFromCache = function () {
                try {
                    var cached = localStorage.getItem('dsp-spr-excel-cache');
                    if (!cached) return;
                    var json = JSON.parse(cached);
                    applySPRFromJson(json, false);
                } catch (e3) { }
            };
            applySPRFromCache();

            document.getElementById('spr-file-btn')?.addEventListener('click', function () {
                document.getElementById('spr-file-input')?.click();
            });

            document.getElementById('spr-file-input')?.addEventListener('change', function (e) {
                var file = e.target.files[0];
                if (!file) return;
                var reader = new FileReader();
                reader.onload = function (ev) {
                    processSPRExcelData(ev.target.result, true);
                };
                reader.readAsArrayBuffer(file);
                this.value = '';
            });

            document.getElementById('adj-reset-btn')?.addEventListener('click', function () {
                const reset = Object.assign({}, SSD_ADJUSTMENT_DEFAULTS);
                saveSSDAdjustments(reset);

                SPR_LIST.forEach(function (ssd) {
                    const el = document.getElementById('adj-' + ssd);
                    if (el) el.value = 0;
                });

                refreshCycleValues();
            });

            document.getElementById('soft-adj-reset-btn')?.addEventListener('click', function () {
                const reset = Object.assign({}, SSD_SOFT_ADJUSTMENT_DEFAULTS);
                saveSSDSoftAdjustments(reset);

                SPR_LIST.forEach(function (ssd) {
                    const el = document.getElementById('soft-adj-' + ssd);
                    if (el) el.value = 0;
                });

                refreshCycleValues();
            });

            ['SSD_C1_1', 'SSD_C1_1B', 'SSD_C3_3B', 'SSD_C3_4'].forEach(function (key) {
                var el = document.getElementById('cvp-' + key);
                if (el) {
                    el.addEventListener('change', function () {
                        var ov = getC1C3Overrides();
                        ov[key] = +this.value || 0;
                        ov['SSD_C1'] = (ov['SSD_C1_1'] || 0) + (ov['SSD_C1_1B'] || 0);
                        ov['SSD_C3'] = (ov['SSD_C3_3B'] || 0) + (ov['SSD_C3_4'] || 0);
                        saveC1C3Overrides(ov);
                        refreshCycleValues();
                    });
                }
            });

            // CVP Buffer イベントリスナー
            var cvpBufMap = [
                { id: 'cvp-buf-1', key: 'CVP_BUF_1' },
                { id: 'cvp-buf-1b', key: 'CVP_BUF_1B' },
                { id: 'cvp-buf-3b', key: 'CVP_BUF_3B' },
                { id: 'cvp-buf-4', key: 'CVP_BUF_4' }
            ];
            cvpBufMap.forEach(function (item) {
                var el = document.getElementById(item.id);
                if (el) {
                    el.addEventListener('change', function () {
                        var ov = getC1C3Overrides();
                        ov[item.key] = +this.value || 0;
                        saveC1C3Overrides(ov);
                        refreshCycleValues();
                    });
                }
            });

            document.getElementById('soft-pct-input')?.addEventListener('change', function () {
                saveSoftPct(Number(this.value) || 0);
                refreshCycleValues();
            });

            document.getElementById('hard-pct-input')?.addEventListener('change', function () {
                saveHardPct(Number(this.value) || 0);
                refreshCycleValues();
            });

            document.getElementById('dl-btn')?.addEventListener('click', downloadExcel);
            document.getElementById('dl-1d-btn')?.addEventListener('click', downloadOneDayExcel);
            document.getElementById('dl-future-req-btn')?.addEventListener('click', downloadFutureRequiredExcel);

            setFutureDownloadButtonState(isFutureRequiredDownloading);

            // 日付変更時にキャッシュからSPR自動反映
            try {
                var cached = localStorage.getItem('dsp-spr-excel-cache');
                if (cached) {
                    var json = JSON.parse(cached);
                    if (json && json.length >= 7) {
                        var headerRow = json[0];
                        var selectedText = getSelectedDateText();
                        var md = parseMonthDayFromText(selectedText);
                        if (md) {
                            var colIdx = -1;
                            for (var c = 0; c < headerRow.length; c++) {
                                var cell = headerRow[c];
                                if (!cell) continue;
                                var cellDate = null;
                                if (typeof cell === 'number') {
                                    cellDate = new Date((cell - 25569) * 86400 * 1000);
                                } else if (typeof cell === 'string') {
                                    cellDate = new Date(cell);
                                }
                                if (cellDate && !isNaN(cellDate.getTime())) {
                                    if (cellDate.getMonth() + 1 === md.month && cellDate.getDate() === md.day) {
                                        colIdx = c;
                                        break;
                                    }
                                }
                            }
                            if (colIdx >= 0) {
                                var mv = getSSDMultipliers();
                                for (var r = 1; r <= 6 && r < json.length; r++) {
                                    var val = json[r][colIdx];
                                    var ssd = SPR_LIST[r - 1];
                                    if (ssd) {
                                        mv[ssd] = (val !== undefined && val !== null && val !== '') ? Math.round(Number(val) * 10) / 10 || 0 : 0;
                                        var el = document.getElementById('mult-' + ssd);
                                        if (el) el.value = mv[ssd].toFixed(1);
                                    }
                                }
                                saveSSDMultipliers(mv);

                                // CVP値も反映 (行11,12,15,16 → 0-indexed: 10,11,14,15)
                                var cvpMap = [
                                    { row: 10, key: 'SSD_C1_1', inputId: 'cvp-SSD_C1_1' },
                                    { row: 11, key: 'SSD_C1_1B', inputId: 'cvp-SSD_C1_1B' },
                                    { row: 14, key: 'SSD_C3_3B', inputId: 'cvp-SSD_C3_3B' },
                                    { row: 15, key: 'SSD_C3_4', inputId: 'cvp-SSD_C3_4' }
                                ];
                                var ov2 = getC1C3Overrides();
                                for (var ci2 = 0; ci2 < cvpMap.length; ci2++) {
                                    var m2 = cvpMap[ci2];
                                    if (m2.row < json.length) {
                                        var cv = json[m2.row][colIdx];
                                        var nv = (cv !== undefined && cv !== null && cv !== '') ? Math.floor(Number(cv)) || 0 : 0;
                                        ov2[m2.key] = nv;
                                        var ce = document.getElementById(m2.inputId);
                                        if (ce) ce.value = nv;
                                    }
                                }
                                ov2['SSD_C1'] = (ov2['SSD_C1_1'] || 0) + (ov2['SSD_C1_1B'] || 0);
                                ov2['SSD_C3'] = (ov2['SSD_C3_3B'] || 0) + (ov2['SSD_C3_4'] || 0);
                                saveC1C3Overrides(ov2);

                                refreshCycleValues();
                            }
                        }
                    }
                }
            } catch (e_cache) { }
        };

        // === 初期化 ===
        const init = () => {
            setTimeout(() => {
                calculateAndDisplay();
                startObserver();
                console.log('[DSP Counter v12.1] 起動完了');
            }, 1000);
        };

        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', init);
        } else {
            init();
        }
    })();
}
