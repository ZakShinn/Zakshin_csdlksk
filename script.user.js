// ==UserScript==
// @name         Hỗ trợ xuất Excel DS CSDL KSK admin.csdlksk.vn
// @namespace    https://hainghia.net/
// @version      1.20.0
// @description  Xuất Excel DS khám sức khỏe + lịch sử sync checkup từ csdlksk
// @match        https://admin.csdlksk.vn/*
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      admin.csdlksk.vn
// @connect      api.emrhub.vn
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @run-at       document-idle
// ==/UserScript==

(function () {
    "use strict";

    // ============================================================
    // CẤU HÌNH
    // ============================================================

    // Hai nguồn API: Nhập liệu + Liên thông (/click)
    // Site dùng cookie phiên (withCredentials) — KHÔNG gửi Bearer
    // (Bearer sai/cũ sẽ bị HTTP 401). Có header phân trang:
    // request: x-request-token ← response: x-next-request-token
    const CHECKUP_REQUEST_OPTS = {
        method: "GET",
        formBody: false,
        usePageFetch: true,
        withCredentials: true,
        // cookie-only: không gắn Authorization
        authMode: "cookie",
        useContinuationToken: true
    };

    const API_SOURCES = [
        {
            label: "Nhập liệu",
            url:
                "https://admin.csdlksk.vn/api/checkup/health-examinations",
            requestOpts: CHECKUP_REQUEST_OPTS
        },
        {
            label: "Liên thông",
            url:
                "https://admin.csdlksk.vn/api/checkup/health-examinations/click",
            requestOpts: CHECKUP_REQUEST_OPTS
        }
    ];

    const HEALTH_DATA_API =
        "https://admin.csdlksk.vn/api/admin/backoffice/reports/sync-checkup/history";

    // Health Data: cookie phiên + page fetch (giống checkup)
    const HEALTH_DATA_REQUEST_OPTS = {
        method: "GET",
        usePageFetch: true,
        withCredentials: true,
        authMode: "cookie",
        useContinuationToken: false
    };

    const PAGE_SIZE = 100;

    // Số request chạy đồng thời.
    // Để 3-5 tương đối an toàn.
    const CONCURRENCY = 4;

    const MAX_RETRY = 3;

    const PAGE_PATHS = {
        checkup:
            "/admin/operation/health-checkup",
        healthData:
            "/admin/operation/health-data"
    };

    // Chuỗi dài hơn mức này bị bỏ (PDF/base64 làm Excel lỗi hoặc hết bộ nhớ).
    // Excel chỉ chứa tối đa 32.767 ký tự/ô.
    const MAX_EXPORT_STRING = 8000;

    const EXCLUDED_FIELD_NAME =
        /base64|pdf_content|file_content|signature_image|raw_file|binary/i;

    // Thứ tự cột quen thuộc. Field mới không có trong danh sách sẽ được thêm sau.
    const PREFERRED_EXPORT_FIELDS = [
        "patient_id",
        "check_up_id",
        "ngay_kham",
        "status",
        "type",
        "times",
        "full_name",
        "id_number",
        "date_of_birth",
        "gender",
        "detail_address",
        "ma_loai_kcb",
        "facility_name",
        "ma_csgd",
        "ten_csgd",
        "ma_csktsk",
        "doctor_name",
        "doctor_citizen_id",
        "id",
        "facility_code",
        "txn_type",
        "tran_id",
        "ten_bn",
        "status_sync",
        "ly_do",
        "created_at",
        "res_message"
    ];

    const FIELD_LABELS = {
        patient_id: "Patient ID",
        check_up_id: "Check Up ID",
        ngay_kham: "Ngày khám",
        status: "Trạng thái",
        type: "Loại",
        times: "Lần khám",
        full_name: "Họ và tên",
        id_number: "CCCD",
        date_of_birth: "Ngày sinh",
        gender: "Giới tính",
        detail_address: "Địa chỉ",
        ma_loai_kcb: "Mã loại KCB",
        facility_name: "Cơ sở",
        ma_csgd: "Mã CSGD",
        ten_csgd: "Tên CSGD",
        ma_csktsk: "Mã CSKTSK",
        doctor_name: "Bác sĩ",
        doctor_citizen_id: "CCCD bác sĩ",
        _nguon: "Nguồn",
        id: "ID",
        facility_code: "Mã CSYT",
        txn_type: "Loại GD",
        tran_id: "Mã giao dịch",
        ten_bn: "Họ và tên BN",
        status_sync: "Trạng thái đồng bộ",
        ly_do: "Lý do",
        created_at: "Thời gian",
        res_message: "Thông báo"
    };

    const droppedExportFields = new Set();

    function isExcludedFieldName(name) {
        return EXCLUDED_FIELD_NAME.test(
            String(name || "")
        );
    }

    function isTextColumnKey(key) {
        return /id_number|citizen|cccd|ma_cs|facility_code|ma_cskb|code$/i
            .test(String(key || ""));
    }

    function rememberDroppedField(name) {
        if (name) {
            droppedExportFields.add(String(name));
        }
    }

    function sanitizeExaminationRow(item) {
        if (!item || typeof item !== "object") {
            return item;
        }

        const cleaned = {};

        for (const [key, value] of Object.entries(item)) {
            if (key.startsWith("_")) {
                cleaned[key] = value;
                continue;
            }

            if (isExcludedFieldName(key)) {
                rememberDroppedField(key);
                continue;
            }

            if (typeof value === "string") {
                if (value.length > MAX_EXPORT_STRING) {
                    rememberDroppedField(key);
                    continue;
                }

                cleaned[key] = value;
                continue;
            }

            if (value && typeof value === "object") {
                let serialized = "";

                try {
                    serialized = JSON.stringify(value);
                } catch {
                    rememberDroppedField(key);
                    continue;
                }

                if (
                    !serialized ||
                    serialized.length > MAX_EXPORT_STRING
                ) {
                    rememberDroppedField(key);
                    continue;
                }

                cleaned[key] = value;
                continue;
            }

            cleaned[key] = value;
        }

        return cleaned;
    }

    // Cắt string quá lớn ngay trên chuỗi response trước JSON.parse
    // để tránh OOM / timeout (PDF base64, field tên lạ nhưng rất dài).
    function stripExcludedFieldsFromJsonText(text) {
        if (!text || typeof text !== "string") {
            return text;
        }

        let output = "";
        let i = 0;

        while (i < text.length) {
            const quote = text.indexOf('"', i);

            if (quote === -1) {
                output += text.slice(i);
                break;
            }

            let j = quote + 1;
            let escaped = false;

            while (j < text.length) {
                const ch = text[j];

                if (escaped) {
                    escaped = false;
                    j++;
                    continue;
                }

                if (ch === "\\") {
                    escaped = true;
                    j++;
                    continue;
                }

                if (ch === '"') {
                    break;
                }

                j++;
            }

            if (j >= text.length) {
                output += text.slice(i);
                break;
            }

            const token = text.slice(quote + 1, j);
            let k = j + 1;

            while (
                k < text.length &&
                /\s/.test(text[k])
            ) {
                k++;
            }

            if (text[k] !== ":") {
                output += text.slice(i, j + 1);
                i = j + 1;
                continue;
            }

            k++;

            while (
                k < text.length &&
                /\s/.test(text[k])
            ) {
                k++;
            }

            if (text[k] !== '"') {
                output += text.slice(i, j + 1);
                i = j + 1;
                continue;
            }

            let v = k + 1;
            let valueEscaped = false;
            let valueLength = 0;

            while (v < text.length) {
                const ch = text[v];

                if (valueEscaped) {
                    valueEscaped = false;
                    v++;
                    valueLength++;
                    continue;
                }

                if (ch === "\\") {
                    valueEscaped = true;
                    v++;
                    continue;
                }

                if (ch === '"') {
                    break;
                }

                v++;
                valueLength++;
            }

            const drop =
                isExcludedFieldName(token) ||
                valueLength > MAX_EXPORT_STRING;

            if (!drop || v >= text.length) {
                output += text.slice(i, j + 1);
                i = j + 1;
                continue;
            }

            rememberDroppedField(token);

            let removeStart = quote;
            let before = quote - 1;

            while (
                before >= i &&
                /\s/.test(text[before])
            ) {
                before--;
            }

            let removeEnd = v + 1;

            if (
                before >= i &&
                text[before] === ","
            ) {
                removeStart = before;
            } else {
                let after = removeEnd;

                while (
                    after < text.length &&
                    /\s/.test(text[after])
                ) {
                    after++;
                }

                if (text[after] === ",") {
                    removeEnd = after + 1;
                }
            }

            output += text.slice(i, removeStart);
            i = removeEnd;
        }

        return output;
    }

    function cleanExportText(value) {
        return String(value)
            .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, "")
            .replace(/[\t\r\n]+/g, " ")
            .replace(/ {2,}/g, " ")
            .trim();
    }

    function formatExportValue(key, value) {
        if (value == null) {
            return "";
        }

        if (key === "gender") {
            if (value === "male") {
                return "Nam";
            }

            if (value === "female") {
                return "Nữ";
            }
        }

        if (
            key === "ngay_kham" ||
            key === "created_at" ||
            /(^ngay_|_at$|datetime)/i.test(key)
        ) {
            return formatDateTime(value);
        }

        if (
            key === "date_of_birth" ||
            /birth|ngay_sinh/i.test(key)
        ) {
            return formatDate(value);
        }

        if (typeof value === "string") {
            const text = cleanExportText(value);

            if (text.length > 32000) {
                return text.slice(0, 32000);
            }

            return text;
        }

        if (typeof value === "number" || typeof value === "boolean") {
            if (isTextColumnKey(key)) {
                return String(value);
            }

            return value;
        }

        if (typeof value === "object") {
            try {
                const serialized = JSON.stringify(value);

                if (
                    !serialized ||
                    serialized.length > MAX_EXPORT_STRING
                ) {
                    rememberDroppedField(key);
                    return "";
                }

                return serialized;
            } catch {
                return "";
            }
        }

        return cleanExportText(value);
    }

    function collectExportFieldKeys(items) {
        const seen = new Set();

        for (const item of items) {
            if (!item || typeof item !== "object") {
                continue;
            }

            for (const key of Object.keys(item)) {
                if (key.startsWith("_") && key !== "_nguon") {
                    continue;
                }

                seen.add(key);
            }
        }

        const ordered = PREFERRED_EXPORT_FIELDS.filter(
            key => seen.has(key)
        );

        const extras = [...seen]
            .filter(key =>
                key !== "_nguon" &&
                !PREFERRED_EXPORT_FIELDS.includes(key)
            )
            .sort();

        ordered.push(...extras);

        if (seen.has("_nguon")) {
            ordered.push("_nguon");
        }

        return ordered;
    }

    function labelForField(key, usedLabels) {
        let label = FIELD_LABELS[key] || key;

        if (usedLabels.has(label)) {
            label = `${label} (${key})`;
        }

        usedLabels.add(label);
        return label;
    }

    function rowsToExcelObjects(items) {
        if (droppedExportFields.size) {
            console.warn(
                "[KSK Export] Bỏ field quá lớn, không đưa vào Excel: " +
                [...droppedExportFields].join(", ")
            );
            droppedExportFields.clear();
        }

        const fieldKeys = collectExportFieldKeys(items);
        const usedLabels = new Set(["STT"]);
        const labels = fieldKeys.map(key =>
            labelForField(key, usedLabels)
        );

        const excelRows = items.map((item, index) => {
            const row = {
                STT: index + 1
            };

            fieldKeys.forEach((key, col) => {
                row[labels[col]] =
                    key === "_nguon"
                        ? (item?._nguon ?? "")
                        : formatExportValue(
                            key,
                            item?.[key]
                        );
            });

            return row;
        });

        return {
            excelRows,
            fieldKeys,
            labels
        };
    }

    function prepareWorksheet(ws, fieldKeys, labels) {
        ws["!cols"] = [
            { wch: 8 },
            ...labels.map((label, index) => {
                const key = fieldKeys[index] || "";
                let width = Math.min(
                    42,
                    Math.max(10, String(label).length + 2)
                );

                if (
                    /address|detail_address|res_message|full_name|facility_name|ten_|doctor_name/i
                        .test(key)
                ) {
                    width = Math.max(width, 28);
                }

                return { wch: width };
            })
        ];

        if (!ws["!ref"]) {
            return;
        }

        ws["!autofilter"] = {
            ref: ws["!ref"]
        };

        const range = XLSX.utils.decode_range(ws["!ref"]);

        for (let c = 1; c <= range.e.c; c++) {
            const key = fieldKeys[c - 1];

            if (!isTextColumnKey(key)) {
                continue;
            }

            for (let r = 1; r <= range.e.r; r++) {
                const address = XLSX.utils.encode_cell({
                    r,
                    c
                });
                const cell = ws[address];

                if (!cell) {
                    continue;
                }

                cell.t = "s";
                cell.z = "@";

                if (cell.v != null) {
                    cell.v = String(cell.v);
                }
            }
        }
    }

    // ============================================================
    // CSS
    // ============================================================

    const style = document.createElement("style");

    style.textContent = `
        #ksk-export-box {
            position: fixed;
            right: 24px;
            bottom: 24px;
            z-index: 999999;
            font-family: Arial, sans-serif;
            display: flex;
            flex-direction: column;
            align-items: flex-end;
            gap: 6px;
        }

        #ksk-export-progress-box {
            display: none;
            width: 300px;
            padding: 12px 14px;
            background: #ffffff;
            border: 1px solid #d9d9d9;
            border-radius: 10px;
            box-shadow: 0 4px 18px rgba(0,0,0,.15);
            color: #333;
            font-size: 13px;
            opacity: 0;
            transform: translateY(14px) scale(0.96);
            transition:
                opacity .38s cubic-bezier(.22, 1, .36, 1),
                transform .38s cubic-bezier(.22, 1, .36, 1),
                box-shadow .38s ease;
            will-change: opacity, transform;
        }

        #ksk-export-progress-box.ksk-visible {
            opacity: 1;
            transform: translateY(0) scale(1);
        }

        #ksk-export-progress-box.ksk-hiding {
            opacity: 0;
            transform: translateY(10px) scale(0.96);
            box-shadow: 0 2px 8px rgba(0,0,0,.08);
            pointer-events: none;
        }

        #ksk-export-progress-text {
            margin-bottom: 8px;
            line-height: 1.4;
        }

        #ksk-export-progress {
            width: 100%;
            height: 9px;
            background: #eee;
            border-radius: 10px;
            overflow: hidden;
        }

        #ksk-export-progress-bar {
            width: 0%;
            height: 100%;
            background: #1677ff;
            transition: width .2s;
        }

        #ksk-export-btn {
            appearance: none;
            border: 0;
            outline: none;
            padding: 11px 18px;
            border-radius: 9px;
            background: #1677ff;
            color: white;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            box-shadow: 0 4px 14px rgba(0,0,0,.18);
        }

        #ksk-export-btn:hover {
            background: #0958d9;
        }

        #ksk-export-btn:disabled {
            background: #8c8c8c;
            cursor: not-allowed;
        }

        #ksk-export-status-success {
            color: #389e0d;
            font-weight: 600;
        }

        #ksk-export-status-error {
            color: #cf1322;
            font-weight: 600;
        }

        #ksk-export-date-box {
            display: none;
            width: 300px;
            padding: 12px 14px;
            background: #ffffff;
            border: 1px solid #d9d9d9;
            border-radius: 10px;
            box-shadow: 0 4px 18px rgba(0,0,0,.15);
            color: #333;
            font-size: 13px;
            box-sizing: border-box;
        }

        #ksk-export-date-box .ksk-date-title {
            font-weight: 600;
            margin-bottom: 8px;
        }

        #ksk-export-date-box .ksk-date-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 8px;
            margin-bottom: 6px;
        }

        #ksk-export-date-box label {
            min-width: 28px;
            color: #555;
        }

        #ksk-export-date-box input[type="date"] {
            flex: 1;
            min-width: 0;
            padding: 4px 6px;
            border: 1px solid #d9d9d9;
            border-radius: 6px;
            font-size: 13px;
        }

        #ksk-export-source-box {
            display: none;
            padding: 3px;
            background: #ffffff;
            border: 1px solid #d9d9d9;
            border-radius: 9px;
            box-shadow: 0 4px 14px rgba(0,0,0,.12);
            box-sizing: border-box;
        }

        #ksk-export-source-box .ksk-source-options {
            display: flex;
            gap: 2px;
        }

        #ksk-export-source-box label {
            position: relative;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0;
            padding: 6px 10px;
            border-radius: 6px;
            cursor: pointer;
            color: #555;
            font-size: 12px;
            font-weight: 600;
            line-height: 1.2;
            white-space: nowrap;
            user-select: none;
        }

        #ksk-export-source-box label:hover {
            color: #1677ff;
            background: #f0f5ff;
        }

        #ksk-export-source-box input[type="radio"] {
            position: absolute;
            opacity: 0;
            pointer-events: none;
        }

        #ksk-export-source-box label:has(input:checked) {
            color: #fff;
            background: #1677ff;
        }

        #ksk-export-source-box label:has(input:checked):hover {
            color: #fff;
            background: #0958d9;
        }
    `;

    document.head.appendChild(style);

    // ============================================================
    // UTIL
    // ============================================================

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    function getCurrentPageMode() {
        const path = location.pathname;

        if (path.startsWith(PAGE_PATHS.checkup)) {
            return "checkup";
        }

        if (path.startsWith(PAGE_PATHS.healthData)) {
            return "healthData";
        }

        return null;
    }

    function isCorrectPage() {
        return getCurrentPageMode() !== null;
    }

    // ============================================================
    // JWT
    // ============================================================

    function decodeJwtPayload(token) {
        try {
            const parts = token.split(".");

            if (parts.length !== 3) {
                return null;
            }

            let payload = parts[1]
                .replace(/-/g, "+")
                .replace(/_/g, "/");

            while (payload.length % 4) {
                payload += "=";
            }

            return JSON.parse(
                decodeURIComponent(
                    atob(payload)
                        .split("")
                        .map(c =>
                            "%" +
                            ("00" + c.charCodeAt(0).toString(16))
                                .slice(-2)
                        )
                        .join("")
                )
            );
        } catch {
            return null;
        }
    }

    function isTokenExpired(token) {
        const payload = decodeJwtPayload(token);

        if (!payload?.exp) {
            return false;
        }

        return Date.now() >= payload.exp * 1000;
    }

    // ============================================================
    // TÌM TOKEN
    // ============================================================

    function searchTokenInValue(value, path = "") {
        if (value === null || value === undefined) {
            return null;
        }

        if (typeof value === "string") {
            const text = value.trim();

            // Bearer eyJ...
            const bearerMatch =
                text.match(/Bearer\s+([A-Za-z0-9._~+/=-]+)/i);

            if (bearerMatch) {
                return {
                    token: bearerMatch[1],
                    path
                };
            }

            // JWT
            const jwtMatch =
                text.match(
                    /eyJ[A-Za-z0-9_-]*\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+/
                );

            if (jwtMatch) {
                return {
                    token: jwtMatch[0],
                    path
                };
            }

            // Có thể dữ liệu là JSON string
            try {
                const parsed = JSON.parse(text);

                return searchTokenInValue(
                    parsed,
                    path
                );
            } catch {
                return null;
            }
        }

        if (typeof value === "object") {
            const priorityKeys = [
                "access_token",
                "accessToken",
                "access-token",
                "token",
                "jwt",
                "id_token",
                "idToken",
                "authorization",
                "authToken",
                "bearerToken"
            ];

            // Tìm các field ưu tiên trước
            for (const wanted of priorityKeys) {
                for (const [key, child] of Object.entries(value)) {
                    if (
                        key.toLowerCase() ===
                        wanted.toLowerCase()
                    ) {
                        const result =
                            searchTokenInValue(
                                child,
                                path
                                    ? `${path}.${key}`
                                    : key
                            );

                        if (result) {
                            return result;
                        }
                    }
                }
            }

            // Sau đó quét toàn bộ object
            for (const [key, child] of Object.entries(value)) {
                const result =
                    searchTokenInValue(
                        child,
                        path
                            ? `${path}.${key}`
                            : key
                    );

                if (result) {
                    return result;
                }
            }
        }

        return null;
    }

    function getPageStorages() {
        try {
            return [
                {
                    name: "localStorage",
                    storage: unsafeWindow.localStorage
                },
                {
                    name: "sessionStorage",
                    storage: unsafeWindow.sessionStorage
                }
            ];
        } catch {
            return [
                {
                    name: "localStorage",
                    storage: localStorage
                },
                {
                    name: "sessionStorage",
                    storage: sessionStorage
                }
            ];
        }
    }

    function findToken() {
        const storages = getPageStorages();

        const priorityWords = [
            "access_token",
            "accesstoken",
            "token",
            "jwt",
            "auth",
            "login",
            "user",
            "account",
            "session"
        ];

        const candidates = [];

        // --------------------------------------------------------
        // ƯU TIÊN KEY CÓ TÊN LIÊN QUAN TOKEN
        // --------------------------------------------------------

        for (const item of storages) {
            const storage = item.storage;

            for (let i = 0; i < storage.length; i++) {
                const key = storage.key(i);

                if (!key) {
                    continue;
                }

                const lower = key.toLowerCase();

                if (
                    !priorityWords.some(word =>
                        lower.includes(word)
                    )
                ) {
                    continue;
                }

                const raw =
                    storage.getItem(key);

                const result =
                    searchTokenInValue(
                        raw,
                        `${item.name}.${key}`
                    );

                if (
                    result?.token &&
                    !isTokenExpired(result.token)
                ) {
                    candidates.push(result);
                }
            }
        }

        // --------------------------------------------------------
        // KHÔNG THẤY -> QUÉT TOÀN BỘ
        // --------------------------------------------------------

        if (!candidates.length) {
            for (const item of storages) {
                const storage = item.storage;

                for (let i = 0; i < storage.length; i++) {
                    const key = storage.key(i);

                    if (!key) {
                        continue;
                    }

                    const raw =
                        storage.getItem(key);

                    const result =
                        searchTokenInValue(
                            raw,
                            `${item.name}.${key}`
                        );

                    if (
                        result?.token &&
                        !isTokenExpired(result.token)
                    ) {
                        candidates.push(result);
                    }
                }
            }
        }

        if (!candidates.length) {
            return null;
        }

        // Ưu tiên token JWT hợp lệ
        const jwt =
            candidates.find(item =>
                decodeJwtPayload(item.token)
            );

        return jwt || candidates[0];
    }

    // ============================================================
    // GM / PAGE REQUEST
    // ============================================================

    function gmRequest(options) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                ...options,

                onload: resolve,

                onerror: reject,

                ontimeout: () =>
                    reject(
                        new Error("Request timeout")
                    )
            });
        });
    }

    async function pageRequest(options) {
        const response =
            await fetch(
                options.url,
                {
                    method:
                        options.method || "GET",

                    credentials:
                        options.withCredentials === false
                            ? "omit"
                            : "include",

                    headers:
                        options.headers || {},

                    body:
                        options.data
                }
            );

        const responseText =
            await response.text();

        const responseHeaders = {};

        try {
            response.headers.forEach(
                (value, key) => {
                    responseHeaders[
                        key.toLowerCase()
                    ] = value;
                }
            );
        } catch {
            // ignore
        }

        return {
            status: response.status,
            responseText,
            responseHeaders
        };
    }

    function getResponseHeader(
        response,
        name
    ) {
        const wanted =
            String(name).toLowerCase();

        if (
            response.responseHeaders &&
            typeof response.responseHeaders ===
                "object" &&
            !Array.isArray(
                response.responseHeaders
            )
        ) {
            return (
                response.responseHeaders[
                    wanted
                ] ||
                null
            );
        }

        const raw =
            response.responseHeaders;

        if (typeof raw !== "string") {
            return null;
        }

        const lines =
            raw.split(/\r?\n/);

        for (const line of lines) {
            const idx =
                line.indexOf(":");

            if (idx === -1) {
                continue;
            }

            const key =
                line
                    .slice(0, idx)
                    .trim()
                    .toLowerCase();

            if (key === wanted) {
                return line
                    .slice(idx + 1)
                    .trim();
            }
        }

        return null;
    }

    // ============================================================
    // FETCH PAGE
    // ============================================================

    async function fetchPage(
        page,
        token,
        apiUrl,
        retry = 0,
        requestOpts = {}
    ) {
        const method =
            requestOpts.method || "GET";

        const formBody =
            requestOpts.formBody === true;

        const extraParams =
            requestOpts.queryParams || {};

        const params = new URLSearchParams();

        params.set(
            "page_number",
            String(page)
        );

        params.set(
            "page_size",
            String(PAGE_SIZE)
        );

        for (const [key, value] of Object.entries(
            extraParams
        )) {
            if (
                value === null ||
                value === undefined ||
                value === ""
            ) {
                continue;
            }

            params.set(
                key,
                String(value)
            );
        }

        const query =
            params.toString();

        // Query trên URL (như Network của trang)
        const url =
            formBody
                ? apiUrl
                : `${apiUrl}?${query}`;

        const usePageFetch =
            requestOpts.usePageFetch === true ||
            (
                typeof location !== "undefined" &&
                apiUrl.startsWith(
                    location.origin
                )
            );

        // cookie = chỉ cookie phiên (admin.csdlksk.vn)
        // bearer = Authorization (khi authMode = bearer)
        const authMode =
            requestOpts.authMode ||
            (
                usePageFetch
                    ? "cookie"
                    : "bearer"
            );

        try {
            const headers = {
                Accept: "application/json"
            };

            if (
                authMode === "bearer" &&
                token
            ) {
                headers.Authorization =
                    `Bearer ${token}`;
            }

            if (
                requestOpts.useContinuationToken &&
                requestOpts.continuationToken &&
                page > 1
            ) {
                headers["x-request-token"] =
                    requestOpts.continuationToken;
            }

            if (formBody) {
                headers["Content-Type"] =
                    "application/x-www-form-urlencoded";
            }

            const requestOptions = {
                method,

                url,

                // Trang có base64_ca rất lớn cần timeout dài hơn
                timeout: 90000,

                headers,

                withCredentials:
                    requestOpts.withCredentials !== false,

                data: formBody
                    ? query
                    : undefined
            };

            const response =
                usePageFetch
                    ? await pageRequest(
                        requestOptions
                    )
                    : await gmRequest(
                        requestOptions
                    );

            if (
                response.status < 200 ||
                response.status >= 300
            ) {
                if (response.status === 401) {
                    throw new Error(
                        "HTTP 401 — phiên đăng nhập hết hạn hoặc thiếu cookie. " +
                        "Hãy F5 trang, đăng nhập lại nếu cần, rồi xuất lại."
                    );
                }

                throw new Error(
                    `HTTP ${response.status}`
                );
            }

            let json;

            try {
                const rawText =
                    stripExcludedFieldsFromJsonText(
                        response.responseText
                    );

                json =
                    JSON.parse(rawText);
            } catch (parseError) {
                throw new Error(
                    "Response không phải JSON / quá lớn: " +
                    (parseError?.message || "")
                );
            }

            if (
                json?.header?.success === false
            ) {
                throw new Error(
                    `${json.header.res_code || ""} ` +
                    `${json.header.res_msg || ""}`
                );
            }

            // Phòng hờ nếu field vẫn còn sau khi parse
            if (
                Array.isArray(
                    json?.data?.data
                )
            ) {
                json.data.data =
                    json.data.data.map(
                        sanitizeExaminationRow
                    );
            }

            const nextToken =
                getResponseHeader(
                    response,
                    "x-next-request-token"
                );

            if (nextToken) {
                json.__continuationToken =
                    nextToken;
            }

            return json;

        } catch (error) {
            if (retry < MAX_RETRY) {
                console.warn(
                    `[KSK Export] Trang ${page} lỗi. ` +
                    `Retry ${retry + 1}/${MAX_RETRY}: ` +
                    `${error?.message || error}`
                );

                await sleep(
                    800 * (retry + 1)
                );

                return fetchPage(
                    page,
                    token,
                    apiUrl,
                    retry + 1,
                    requestOpts
                );
            }

            throw error;
        }
    }

    // ============================================================
    // DATE FORMAT
    // ============================================================

    function formatDate(value) {
        if (!value) {
            return "";
        }

        const str =
            String(value);

        if (!/^\d{8}/.test(str)) {
            return str;
        }

        const yyyy =
            str.slice(0, 4);

        const mm =
            str.slice(4, 6);

        const dd =
            str.slice(6, 8);

        return `${dd}/${mm}/${yyyy}`;
    }

    function pad2(n) {
        return String(n).padStart(2, "0");
    }

    function formatDateTime(value) {
        if (!value) {
            return "";
        }

        const str =
            String(value);

        if (/^\d{12}/.test(str)) {
            const yyyy =
                str.slice(0, 4);

            const mm =
                str.slice(4, 6);

            const dd =
                str.slice(6, 8);

            const hh =
                str.slice(8, 10);

            const mi =
                str.slice(10, 12);

            return (
                `${dd}/${mm}/${yyyy} ` +
                `${hh}:${mi}`
            );
        }

        // ISO: 2026-09-04T17:15:32.317196
        if (
            /^\d{4}-\d{2}-\d{2}T/.test(str)
        ) {
            const d = new Date(str);

            if (!Number.isNaN(d.getTime())) {
                return (
                    `${pad2(d.getDate())}/` +
                    `${pad2(d.getMonth() + 1)}/` +
                    `${d.getFullYear()} ` +
                    `${pad2(d.getHours())}:` +
                    `${pad2(d.getMinutes())}:` +
                    `${pad2(d.getSeconds())}`
                );
            }
        }

        return formatDate(str);
    }

    function buildExportFileName(prefix) {
        const now = new Date();

        return (
            `${prefix}_` +
            `${now.getFullYear()}-` +
            `${pad2(now.getMonth() + 1)}-` +
            `${pad2(now.getDate())}_` +
            `${pad2(now.getHours())}-` +
            `${pad2(now.getMinutes())}.xlsx`
        );
    }

    function toInputDateValue(date) {
        return (
            `${date.getFullYear()}-` +
            `${pad2(date.getMonth() + 1)}-` +
            `${pad2(date.getDate())}`
        );
    }

    function getDefaultHealthDataRange() {
        const end = new Date();
        const start = new Date(
            end.getFullYear(),
            end.getMonth(),
            1
        );

        return {
            start: toInputDateValue(start),
            end: toInputDateValue(end)
        };
    }

    // API: YYYYMMDDHHmmss
    function toApiDateTime(
        inputDate,
        endOfDay = false
    ) {
        if (!inputDate) {
            return "";
        }

        const parts =
            String(inputDate).split("-");

        if (parts.length !== 3) {
            return "";
        }

        const [yyyy, mm, dd] = parts;

        if (endOfDay) {
            return `${yyyy}${mm}${dd}235959`;
        }

        return `${yyyy}${mm}${dd}000000`;
    }

    function getHealthDataDateRange() {
        const startInput =
            document.getElementById(
                "ksk-export-start-date"
            );

        const endInput =
            document.getElementById(
                "ksk-export-end-date"
            );

        const defaults =
            getDefaultHealthDataRange();

        const startValue =
            startInput?.value ||
            defaults.start;

        const endValue =
            endInput?.value ||
            defaults.end;

        if (startValue > endValue) {
            throw new Error(
                "Ngày bắt đầu phải nhỏ hơn hoặc bằng ngày kết thúc."
            );
        }

        return {
            startInput: startValue,
            endInput: endValue,
            start_date: toApiDateTime(
                startValue,
                false
            ),
            end_date: toApiDateTime(
                endValue,
                true
            )
        };
    }

    // ============================================================
    // UI
    // ============================================================

    let exportBox;
    let exportButton;
    let progressBox;
    let progressText;
    let progressBar;
    let dateBox;
    let sourceBox;
    let hideProgressTimer = null;
    let hideProgressAnimTimer = null;

    const PROGRESS_HIDE_DELAY_MS = 10000;
    const PROGRESS_HIDE_ANIM_MS = 380;

    function createUI() {
        if (
            document.getElementById(
                "ksk-export-box"
            )
        ) {
            syncFilterUI();
            return;
        }

        exportBox =
            document.createElement("div");

        exportBox.id =
            "ksk-export-box";

        dateBox =
            document.createElement("div");

        dateBox.id =
            "ksk-export-date-box";

        const defaults =
            getDefaultHealthDataRange();

        dateBox.innerHTML = `
            <div class="ksk-date-title">Khoảng thời gian xuất</div>
            <div class="ksk-date-row">
                <label for="ksk-export-start-date">Từ</label>
                <input type="date" id="ksk-export-start-date" value="${defaults.start}">
            </div>
            <div class="ksk-date-row">
                <label for="ksk-export-end-date">Đến</label>
                <input type="date" id="ksk-export-end-date" value="${defaults.end}">
            </div>
        `;

        sourceBox =
            document.createElement("div");

        sourceBox.id =
            "ksk-export-source-box";

        sourceBox.innerHTML = `
            <div class="ksk-source-options">
                <label>
                    <input type="radio" name="ksk-export-source" value="Nhập liệu">
                    Nhập liệu
                </label>
                <label>
                    <input type="radio" name="ksk-export-source" value="Liên thông" checked>
                    Liên thông
                </label>
                <label>
                    <input type="radio" name="ksk-export-source" value="all">
                    Cả 2
                </label>
            </div>
        `;

        progressBox =
            document.createElement("div");

        progressBox.id =
            "ksk-export-progress-box";

        progressText =
            document.createElement("div");

        progressText.id =
            "ksk-export-progress-text";

        progressText.textContent =
            "Sẵn sàng.";

        const progress =
            document.createElement("div");

        progress.id =
            "ksk-export-progress";

        progressBar =
            document.createElement("div");

        progressBar.id =
            "ksk-export-progress-bar";

        progress.appendChild(
            progressBar
        );

        progressBox.appendChild(
            progressText
        );

        progressBox.appendChild(
            progress
        );

        exportButton =
            document.createElement("button");

        exportButton.id =
            "ksk-export-btn";

        exportButton.innerHTML =
            "📊 Xuất Excel";

        exportButton.addEventListener(
            "click",
            onExportClick
        );

        exportBox.appendChild(
            dateBox
        );

        exportBox.appendChild(
            sourceBox
        );

        exportBox.appendChild(
            progressBox
        );

        exportBox.appendChild(
            exportButton
        );

        document.body.appendChild(
            exportBox
        );

        syncFilterUI();
    }

    function syncFilterUI() {
        const mode = getCurrentPageMode();

        const dateFilter =
            dateBox ||
            document.getElementById(
                "ksk-export-date-box"
            );

        if (dateFilter) {
            dateFilter.style.display =
                mode === "healthData"
                    ? "block"
                    : "none";
        }

        const sourceFilter =
            sourceBox ||
            document.getElementById(
                "ksk-export-source-box"
            );

        if (sourceFilter) {
            sourceFilter.style.display =
                mode === "checkup"
                    ? "block"
                    : "none";
        }
    }

    function getSelectedCheckupSources() {
        const selected =
            document.querySelector(
                'input[name="ksk-export-source"]:checked'
            )?.value;

        if (selected === "all") {
            return API_SOURCES.slice();
        }

        const label =
            selected === "Nhập liệu"
                ? "Nhập liệu"
                : "Liên thông";

        const sources =
            API_SOURCES.filter(
                source =>
                    source.label === label
            );

        if (sources.length) {
            return sources;
        }

        return API_SOURCES.filter(
            source =>
                source.label === "Liên thông"
        );
    }

    function removeUI() {
        document
            .getElementById(
                "ksk-export-box"
            )
            ?.remove();

        exportBox = null;
        exportButton = null;
        progressBox = null;
        progressText = null;
        progressBar = null;
        dateBox = null;
        sourceBox = null;
        clearHideProgressTimers();
    }

    function clearHideProgressTimers() {
        if (hideProgressTimer) {
            clearTimeout(hideProgressTimer);
            hideProgressTimer = null;
        }

        if (hideProgressAnimTimer) {
            clearTimeout(hideProgressAnimTimer);
            hideProgressAnimTimer = null;
        }
    }

    function showProgressBox() {
        if (!progressBox) {
            return;
        }

        clearHideProgressTimers();

        progressBox.classList.remove(
            "ksk-hiding"
        );

        const alreadyOpen =
            progressBox.style.display === "block" &&
            progressBox.classList.contains(
                "ksk-visible"
            );

        if (alreadyOpen) {
            return;
        }

        progressBox.style.display =
            "block";

        // Force reflow rồi mới add class để chạy animation
        void progressBox.offsetWidth;

        requestAnimationFrame(() => {
            progressBox?.classList.add(
                "ksk-visible"
            );
        });
    }

    function hideProgressBox() {
        if (!progressBox) {
            return;
        }

        clearHideProgressTimers();

        progressBox.classList.remove(
            "ksk-visible"
        );

        progressBox.classList.add(
            "ksk-hiding"
        );

        hideProgressAnimTimer =
            setTimeout(() => {
                if (!progressBox) {
                    return;
                }

                progressBox.style.display =
                    "none";

                progressBox.classList.remove(
                    "ksk-hiding"
                );

                hideProgressAnimTimer =
                    null;
            }, PROGRESS_HIDE_ANIM_MS);
    }

    function scheduleHideProgress(
        delayMs = PROGRESS_HIDE_DELAY_MS
    ) {
        clearHideProgressTimers();

        hideProgressTimer =
            setTimeout(() => {
                hideProgressTimer = null;
                hideProgressBox();
            }, delayMs);
    }

    function setProgress(
        percent,
        text,
        options = {}
    ) {
        showProgressBox();

        if (progressBar) {
            progressBar.style.width =
                `${Math.min(100, percent)}%`;
        }

        if (progressText) {
            progressText.innerHTML =
                text;
        }

        if (options.autoHide) {
            scheduleHideProgress(
                options.hideAfterMs ??
                PROGRESS_HIDE_DELAY_MS
            );
        }
    }

    // ============================================================
    // CHẠY NHIỀU REQUEST SONG SONG
    // ============================================================

    async function fetchPagesConcurrent(
        pageNumbers,
        token,
        apiUrl,
        onPageDone,
        requestOpts = {}
    ) {
        const results = [];

        let index = 0;

        async function worker() {
            while (true) {
                const current =
                    index++;

                if (
                    current >=
                    pageNumbers.length
                ) {
                    break;
                }

                const page =
                    pageNumbers[current];

                try {
                    const data =
                        await fetchPage(
                            page,
                            token,
                            apiUrl,
                            0,
                            requestOpts
                        );

                    results.push({
                        page,
                        success: true,
                        data
                    });

                } catch (error) {
                    results.push({
                        page,
                        success: false,
                        error
                    });
                }

                onPageDone?.(
                    page,
                    results.length
                );
            }
        }

        const workers =
            Array.from(
                {
                    length:
                        Math.min(
                            CONCURRENCY,
                            pageNumbers.length
                        )
                },
                () => worker()
            );

        await Promise.all(workers);

        return results;
    }

    function tagRowsWithSource(
        rows,
        sourceLabel
    ) {
        return rows.map(item => ({
            ...sanitizeExaminationRow(
                item
            ),
            _nguon: sourceLabel
        }));
    }

    async function fetchAllFromSource(
        source,
        token,
        onProgress
    ) {
        const requestOpts =
            source.requestOpts || {};

        const zeroBased =
            source.zeroBasedPages === true;

        const firstPageNum =
            zeroBased ? 0 : 1;

        const first =
            await fetchPage(
                firstPageNum,
                token,
                source.url,
                0,
                requestOpts
            );

        const firstData =
            first?.data;

        if (
            !firstData ||
            !Array.isArray(
                firstData.data
            )
        ) {
            throw new Error(
                `Response API [${source.label}] không đúng cấu trúc.`
            );
        }

        const totalPages =
            Number(
                firstData.total_pages
            ) || 1;

        const totalElements =
            Number(
                firstData.total_elements
            ) || 0;

        let rows =
            tagRowsWithSource(
                firstData.data,
                source.label
            );

        console.log(
            `[KSK Export] [${source.label}] ` +
            `Tổng hồ sơ: ${totalElements}, ` +
            `trang: ${totalPages}` +
            (
                zeroBased
                    ? " (page 0-based)"
                    : ""
            )
        );

        onProgress?.(
            1,
            totalPages,
            source.label
        );

        if (totalPages > 1) {
            const lastPageNum =
                zeroBased
                    ? totalPages - 1
                    : totalPages;

            const useContinuation =
                requestOpts.useContinuationToken === true;

            // Phân trang có x-request-token phải tuần tự
            if (useContinuation) {
                let continuationToken =
                    first.__continuationToken ||
                    null;

                let completed = 1;

                for (
                    let page = firstPageNum + 1;
                    page <= lastPageNum;
                    page++
                ) {
                    const pageOpts = {
                        ...requestOpts,
                        continuationToken
                    };

                    const pageJson =
                        await fetchPage(
                            page,
                            token,
                            source.url,
                            0,
                            pageOpts
                        );

                    const pageRows =
                        pageJson
                            ?.data
                            ?.data;

                    if (
                        Array.isArray(
                            pageRows
                        )
                    ) {
                        rows.push(
                            ...tagRowsWithSource(
                                pageRows,
                                source.label
                            )
                        );
                    }

                    if (
                        pageJson.__continuationToken
                    ) {
                        continuationToken =
                            pageJson.__continuationToken;
                    }

                    completed++;

                    onProgress?.(
                        completed,
                        totalPages,
                        source.label
                    );
                }

            } else {
                const pages = [];

                for (
                    let page = firstPageNum + 1;
                    page <= lastPageNum;
                    page++
                ) {
                    pages.push(page);
                }

                let completed = 1;

                const results =
                    await fetchPagesConcurrent(
                        pages,
                        token,
                        source.url,
                        () => {
                            completed++;

                            onProgress?.(
                                completed,
                                totalPages,
                                source.label
                            );
                        },
                        requestOpts
                    );

                const failed = [];

                results
                    .sort(
                        (a, b) =>
                            a.page - b.page
                    )
                    .forEach(result => {
                        if (
                            result.success
                        ) {
                            const pageRows =
                                result
                                    .data
                                    ?.data
                                    ?.data;

                            if (
                                Array.isArray(
                                    pageRows
                                )
                            ) {
                                rows.push(
                                    ...tagRowsWithSource(
                                        pageRows,
                                        source.label
                                    )
                                );
                            }

                        } else {
                            failed.push(
                                result.page
                            );
                        }
                    });

                if (failed.length) {
                    throw new Error(
                        `[${source.label}] Không tải được trang: ` +
                        failed.join(", ")
                    );
                }
            }
        }

        return {
            label: source.label,
            rows,
            totalElements,
            totalPages
        };
    }

    // ============================================================
    // EXPORT
    // ============================================================

    async function onExportClick() {
        const mode = getCurrentPageMode();

        if (mode === "checkup") {
            await exportCheckup();
            return;
        }

        if (mode === "healthData") {
            await exportHealthDataHistory();
            return;
        }

        alert(
            "Chức năng xuất Excel chỉ chạy tại trang Health Checkup hoặc Health Data."
        );
    }

    async function exportHealthDataHistory() {
        exportButton.disabled = true;
        exportButton.textContent =
            "⏳ Đang xử lý...";

        try {
            const dateRange =
                getHealthDataDateRange();

            setProgress(
                2,
                "🔐 Đang kiểm tra phiên đăng nhập..."
            );

            // API admin.csdlksk.vn auth bằng cookie phiên.
            const token = null;

            console.log(
                "[KSK Export][Health Data] Auth: cookie session"
            );

            console.log(
                "[KSK Export][Health Data] Range:",
                dateRange.start_date,
                "→",
                dateRange.end_date
            );

            setProgress(
                5,
                `📥 Đang tải ${dateRange.startInput} → ${dateRange.endInput}...`
            );

            const result =
                await fetchAllFromSource(
                    {
                        label: "Health Data",
                        url: HEALTH_DATA_API,
                        zeroBasedPages: true,
                        requestOpts: {
                            ...HEALTH_DATA_REQUEST_OPTS,
                            queryParams: {
                                start_date:
                                    dateRange.start_date,
                                end_date:
                                    dateRange.end_date
                            }
                        }
                    },
                    token,
                    (
                        completed,
                        totalPages
                    ) => {
                        const percent =
                            Math.round(
                                5 +
                                (
                                    completed /
                                    totalPages
                                ) *
                                85
                            );

                        setProgress(
                            percent,
                            `📥 Đang tải ${dateRange.startInput} → ${dateRange.endInput}<br>` +
                            `${completed}/${totalPages} trang`
                        );
                    }
                );

            const allData =
                result.rows.map(
                    ({ _nguon, ...rest }) =>
                        rest
                );

            setProgress(
                92,
                "🧹 Đang xử lý dữ liệu..."
            );

            const unique =
                new Map();

            for (const item of allData) {
                const key =
                    item.id ??
                    (
                        `${item.facility_code ?? ""}_` +
                        `${item.ten_bn ?? ""}_` +
                        `${item.created_at ?? ""}`
                    );

                unique.set(
                    String(key),
                    item
                );
            }

            const rows =
                [...unique.values()];

            const sheet =
                rowsToExcelObjects(rows);

            const excelRows =
                sheet.excelRows;

            setProgress(
                95,
                "📊 Đang tạo file Excel..."
            );

            const ws =
                XLSX.utils.json_to_sheet(
                    excelRows
                );

            prepareWorksheet(
                ws,
                sheet.fieldKeys,
                sheet.labels
            );

            const wb =
                XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(
                wb,
                ws,
                "Lịch sử sync"
            );

            const fileName =
                buildExportFileName(
                    "csdlksk_health_data"
                );

            XLSX.writeFile(
                wb,
                fileName
            );

            setProgress(
                100,
                `<span id="ksk-export-status-success">` +
                `✅ Hoàn tất: ${excelRows.length} bản ghi` +
                `</span>`,
                { autoHide: true }
            );

            exportButton.textContent =
                "✅ Xuất Excel";

            console.log(
                `[KSK Export][Health Data] Hoàn tất ${excelRows.length} bản ghi`
            );

            if (
                result.totalElements &&
                excelRows.length !==
                    result.totalElements
            ) {
                console.warn(
                    `[KSK Export][Health Data] API báo ${result.totalElements}, ` +
                    `file có ${excelRows.length} bản ghi.`
                );
            }

            setTimeout(() => {
                exportButton.textContent =
                    "📊 Xuất Excel";
            }, 2500);

        } catch (error) {
            console.error(
                "[KSK Export][Health Data]",
                error
            );

            setProgress(
                0,
                `<span id="ksk-export-status-error">` +
                `❌ ${error.message}` +
                `</span>`,
                { autoHide: true }
            );

            exportButton.textContent =
                "❌ Xuất lỗi";

            setTimeout(() => {
                exportButton.textContent =
                    "📊 Xuất Excel";
            }, 3000);

        } finally {
            exportButton.disabled =
                false;
        }
    }

    async function exportCheckup() {
        if (getCurrentPageMode() !== "checkup") {
            alert(
                "Chức năng này chỉ chạy tại trang Health Checkup."
            );

            return;
        }

        exportButton.disabled =
            true;

        exportButton.textContent =
            "⏳ Đang xử lý...";

        try {
            // ----------------------------------------------------
            // PHIÊN ĐĂNG NHẬP (cookie) — không dùng Bearer
            // ----------------------------------------------------

            setProgress(
                2,
                "🔐 Đang kiểm tra phiên đăng nhập..."
            );

            // API admin.csdlksk.vn auth bằng cookie phiên.
            // Gửi Bearer sai/cũ sẽ bị 401 — nên không dùng token storage.
            const token = null;

            console.log(
                "[KSK Export] Auth: cookie session"
            );

            // ----------------------------------------------------
            // TẢI NGUỒN ĐÃ CHỌN: Nhập liệu hoặc Liên thông
            // ----------------------------------------------------

            let allData = [];
            let totalElementsExpected = 0;

            const selectedSources =
                getSelectedCheckupSources();

            console.log(
                "[KSK Export] Nguồn xuất:",
                selectedSources
                    .map(s => s.label)
                    .join(", ")
            );

            const sourceCount =
                selectedSources.length;

            for (
                let sourceIndex = 0;
                sourceIndex < sourceCount;
                sourceIndex++
            ) {
                const source =
                    selectedSources[
                        sourceIndex
                    ];

                const progressBase =
                    5 +
                    (
                        sourceIndex /
                        sourceCount
                    ) *
                    85;

                const progressSpan =
                    85 / sourceCount;

                setProgress(
                    Math.round(
                        progressBase
                    ),
                    `📥 Đang tải [${source.label}]...`
                );

                const result =
                    await fetchAllFromSource(
                        source,
                        token,
                        (
                            completed,
                            totalPages,
                            label
                        ) => {
                            const percent =
                                Math.round(
                                    progressBase +
                                    (
                                        completed /
                                        totalPages
                                    ) *
                                    progressSpan
                                );

                            setProgress(
                                percent,
                                `📥 [${label}] Đang tải...<br>` +
                                `${completed}/${totalPages} trang`
                            );
                        }
                    );

                allData.push(
                    ...result.rows
                );

                totalElementsExpected +=
                    result.totalElements;

                console.log(
                    `[KSK Export] [${result.label}] ` +
                    `Đã lấy ${result.rows.length} hồ sơ`
                );
            }

            // ----------------------------------------------------
            // LOẠI TRÙNG (trong cùng nguồn)
            // ----------------------------------------------------

            setProgress(
                90,
                "🧹 Đang xử lý dữ liệu..."
            );

            const unique =
                new Map();

            for (const item of allData) {
                const baseKey =
                    item.check_up_id ??
                    (
                        `${item.patient_id ?? ""}_` +
                        `${item.ngay_kham ?? ""}_` +
                        `${item.id_number ?? ""}`
                    );

                // Giữ cả hai nguồn nếu cùng hồ sơ xuất hiện ở cả hai API
                const key =
                    `${item._nguon || ""}|${baseKey}`;

                unique.set(
                    String(key),
                    item
                );
            }

            allData =
                [...unique.values()];

            // ----------------------------------------------------
            // EXCEL ROWS
            // ----------------------------------------------------

            const sheet =
                rowsToExcelObjects(allData);

            const excelRows =
                sheet.excelRows;

            // ----------------------------------------------------
            // XLSX
            // ----------------------------------------------------

            setProgress(
                95,
                "📊 Đang tạo file Excel..."
            );

            const ws =
                XLSX.utils.json_to_sheet(
                    excelRows
                );

            prepareWorksheet(
                ws,
                sheet.fieldKeys,
                sheet.labels
            );

            const wb =
                XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(
                wb,
                ws,
                "Danh sách khám"
            );

            // ----------------------------------------------------
            // TÊN FILE
            // ----------------------------------------------------

            const fileName =
                buildExportFileName(
                    "csdlksk"
                );

            XLSX.writeFile(
                wb,
                fileName
            );

            // ----------------------------------------------------
            // DONE
            // ----------------------------------------------------

            setProgress(
                100,
                `<span id="ksk-export-status-success">` +
                `✅ Hoàn tất: ${excelRows.length} hồ sơ` +
                `</span>`,
                { autoHide: true }
            );

            exportButton.textContent =
                "✅ Xuất Excel";

            console.log(
                `[KSK Export] Hoàn tất ${excelRows.length} hồ sơ`
            );

            if (
                totalElementsExpected &&
                excelRows.length !==
                    totalElementsExpected
            ) {
                console.warn(
                    `[KSK Export] API báo ${totalElementsExpected} hồ sơ, ` +
                    `file có ${excelRows.length} hồ sơ.`
                );
            }

            setTimeout(() => {
                exportButton.textContent =
                    "📊 Xuất Excel";
            }, 2500);

        } catch (error) {
            console.error(
                "[KSK Export]",
                error
            );

            setProgress(
                0,
                `<span id="ksk-export-status-error">` +
                `❌ ${error.message}` +
                `</span>`,
                { autoHide: true }
            );

            exportButton.textContent =
                "❌ Xuất lỗi";

            setTimeout(() => {
                exportButton.textContent =
                    "📊 Xuất Excel";
            }, 3000);

        } finally {
            exportButton.disabled =
                false;
        }
    }

    // ============================================================
    // HỖ TRỢ WEBSITE SPA
    // ============================================================

    function updateUI() {
        if (isCorrectPage()) {
            createUI();
            syncFilterUI();
        } else {
            removeUI();
        }
    }

    let lastUrl =
        location.href;

    let updateTimer = null;

    function scheduleUpdateUI() {
        if (updateTimer) {
            clearTimeout(updateTimer);
        }

        updateTimer = setTimeout(() => {
            updateTimer = null;
            updateUI();
        }, 200);
    }

    function onUrlMaybeChanged() {
        const current =
            location.href;

        if (current === lastUrl) {
            return;
        }

        lastUrl = current;
        scheduleUpdateUI();
    }

    // Hook History API (SPA navigate không reload trang)
    const historyMethods = [
        "pushState",
        "replaceState"
    ];

    for (const method of historyMethods) {
        const original =
            history[method];

        if (typeof original !== "function") {
            continue;
        }

        history[method] = function (...args) {
            const result =
                original.apply(
                    this,
                    args
                );

            onUrlMaybeChanged();
            return result;
        };
    }

    window.addEventListener(
        "popstate",
        onUrlMaybeChanged
    );

    // Một số router SPA dùng hash
    window.addEventListener(
        "hashchange",
        onUrlMaybeChanged
    );

    // Theo dõi DOM: framework có thể xóa nút, hoặc đổi URL ngoài History API
    const observer =
        new MutationObserver(() => {
            onUrlMaybeChanged();

            if (
                isCorrectPage() &&
                !document.getElementById(
                    "ksk-export-box"
                )
            ) {
                createUI();
            }
        });

    observer.observe(
        document.documentElement,
        {
            childList: true,
            subtree: true
        }
    );

    // Load ban đầu + vài lần sau (SPA render muộn)
    updateUI();
    setTimeout(updateUI, 500);
    setTimeout(updateUI, 1500);
    setTimeout(updateUI, 3000);

})();
