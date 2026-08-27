// ==UserScript==
// @name         Hỗ trợ xuất Excel DS CSDL KSK admin.csdlksk.vn
// @namespace    https://hainghia.net/
// @version      1.10.2
// @description  Xuất toàn bộ danh sách khám sức khỏe từ csdlksk ra Excel
// @match        https://admin.csdlksk.vn/*
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
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
    const API_SOURCES = [
        {
            label: "Nhập liệu",
            url:
                "https://api.emrhub.vn/api/checkup/health-examinations"
        },
        {
            label: "Liên thông",
            url:
                "https://api.emrhub.vn/api/checkup/health-examinations/click"
        }
    ];

    const PAGE_SIZE = 100;

    // Số request chạy đồng thời.
    // Để 3-5 tương đối an toàn.
    const CONCURRENCY = 4;

    const MAX_RETRY = 3;

    const TARGET_PATH =
        "/admin/operation/health-checkup";

    // Field API không dùng khi xuất Excel (vd. PDF base64 rất lớn)
    const EXCLUDED_EXPORT_FIELDS = [
        "base64_ca"
    ];

    function sanitizeExaminationRow(item) {
        if (!item || typeof item !== "object") {
            return item;
        }

        const cleaned = { ...item };

        for (const field of EXCLUDED_EXPORT_FIELDS) {
            delete cleaned[field];
        }

        return cleaned;
    }

    // Cắt field lớn (PDF base64) ngay trên chuỗi response trước JSON.parse
    // để tránh OOM / timeout khi trang có hồ sơ completed.
    function stripExcludedFieldsFromJsonText(text) {
        if (!text || typeof text !== "string") {
            return text;
        }

        let result = text;

        for (const field of EXCLUDED_EXPORT_FIELDS) {
            const key = `"${field}"`;
            let output = "";
            let i = 0;

            while (true) {
                const start =
                    result.indexOf(key, i);

                if (start === -1) {
                    output +=
                        result.slice(i);
                    break;
                }

                let j =
                    start + key.length;

                while (
                    j < result.length &&
                    /\s/.test(result[j])
                ) {
                    j++;
                }

                if (result[j] !== ":") {
                    output +=
                        result.slice(
                            i,
                            start + key.length
                        );
                    i =
                        start + key.length;
                    continue;
                }

                j++;

                while (
                    j < result.length &&
                    /\s/.test(result[j])
                ) {
                    j++;
                }

                if (result[j] !== '"') {
                    // Không phải string → bỏ qua
                    output +=
                        result.slice(
                            i,
                            start + key.length
                        );
                    i =
                        start + key.length;
                    continue;
                }

                j++;

                while (
                    j < result.length &&
                    result[j] !== '"'
                ) {
                    // Base64 thường không escape; vẫn bỏ qua \"
                    if (
                        result[j] === "\\" &&
                        j + 1 < result.length
                    ) {
                        j += 2;
                        continue;
                    }

                    j++;
                }

                if (j >= result.length) {
                    output +=
                        result.slice(i);
                    break;
                }

                j++;

                let removeStart = start;
                let k = start - 1;

                while (
                    k >= i &&
                    /\s/.test(result[k])
                ) {
                    k--;
                }

                if (
                    k >= i &&
                    result[k] === ","
                ) {
                    removeStart = k;
                } else {
                    let m = j;

                    while (
                        m < result.length &&
                        /\s/.test(result[m])
                    ) {
                        m++;
                    }

                    if (result[m] === ",") {
                        j = m + 1;
                    }
                }

                output +=
                    result.slice(
                        i,
                        removeStart
                    );
                i = j;
            }

            result = output;
        }

        return result;
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
            gap: 8px;
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
    `;

    document.head.appendChild(style);

    // ============================================================
    // UTIL
    // ============================================================

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    function isCorrectPage() {
        return location.pathname.startsWith(TARGET_PATH);
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
    // GM REQUEST
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

    // ============================================================
    // FETCH PAGE
    // ============================================================

    async function fetchPage(
        page,
        token,
        apiUrl,
        retry = 0
    ) {
        const url =
            `${apiUrl}` +
            `?page_number=${page}` +
            `&page_size=${PAGE_SIZE}`;

        try {
            const response =
                await gmRequest({
                    method: "GET",

                    url,

                    // Trang có base64_ca rất lớn cần timeout dài hơn
                    timeout: 90000,

                    headers: {
                        Accept: "application/json",

                        Authorization:
                            `Bearer ${token}`
                    }
                });

            if (
                response.status < 200 ||
                response.status >= 300
            ) {
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
                    retry + 1
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

        return formatDate(str);
    }

    // ============================================================
    // UI
    // ============================================================

    let exportBox;
    let exportButton;
    let progressBox;
    let progressText;
    let progressBar;

    function createUI() {
        if (
            document.getElementById(
                "ksk-export-box"
            )
        ) {
            return;
        }

        exportBox =
            document.createElement("div");

        exportBox.id =
            "ksk-export-box";

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
            exportAll
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
    }

    function removeUI() {
        document
            .getElementById(
                "ksk-export-box"
            )
            ?.remove();
    }

    function setProgress(
        percent,
        text
    ) {
        if (progressBox) {
            progressBox.style.display =
                "block";
        }

        if (progressBar) {
            progressBar.style.width =
                `${Math.min(100, percent)}%`;
        }

        if (progressText) {
            progressText.innerHTML =
                text;
        }
    }

    // ============================================================
    // CHẠY NHIỀU REQUEST SONG SONG
    // ============================================================

    async function fetchPagesConcurrent(
        pageNumbers,
        token,
        apiUrl,
        onPageDone
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
                            apiUrl
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
        const first =
            await fetchPage(
                1,
                token,
                source.url
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
            `trang: ${totalPages}`
        );

        onProgress?.(
            1,
            totalPages,
            source.label
        );

        if (totalPages > 1) {
            const pages = [];

            for (
                let page = 2;
                page <= totalPages;
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
                    }
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

    async function exportAll() {
        if (!isCorrectPage()) {
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
            // TOKEN
            // ----------------------------------------------------

            setProgress(
                2,
                "🔑 Đang tìm token đăng nhập..."
            );

            const tokenInfo =
                findToken();

            if (!tokenInfo?.token) {
                throw new Error(
                    "Không tìm thấy access token. " +
                    "Hãy đăng nhập lại trang rồi thử lại."
                );
            }

            const token =
                tokenInfo.token;

            console.log(
                "[KSK Export] Token:",
                tokenInfo.path
            );

            // ----------------------------------------------------
            // TẢI CẢ HAI NGUỒN: Nhập liệu + Liên thông
            // ----------------------------------------------------

            let allData = [];
            let totalElementsExpected = 0;

            const sourceCount =
                API_SOURCES.length;

            for (
                let sourceIndex = 0;
                sourceIndex < sourceCount;
                sourceIndex++
            ) {
                const source =
                    API_SOURCES[
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

            const excelRows =
                allData.map(
                    (item, index) => ({
                        "STT":
                            index + 1,

                        "Patient ID":
                            item.patient_id ?? "",

                        "Check Up ID":
                            item.check_up_id ?? "",

                        "Ngày khám":
                            formatDateTime(
                                item.ngay_kham
                            ),

                        "Trạng thái":
                            item.status ?? "",

                        "Loại":
                            item.type ?? "",

                        "Lần khám":
                            item.times ?? "",

                        "Họ và tên":
                            item.full_name ?? "",

                        "CCCD":
                            item.id_number != null
                                ? String(
                                    item.id_number
                                )
                                : "",

                        "Ngày sinh":
                            formatDate(
                                item.date_of_birth
                            ),

                        "Giới tính":
                            item.gender === "male"
                                ? "Nam"
                                : item.gender === "female"
                                    ? "Nữ"
                                    : (
                                        item.gender ?? ""
                                    ),

                        "Địa chỉ":
                            item.detail_address ?? "",

                        "Nguồn":
                            item._nguon ?? ""
                    })
                );

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

            ws["!cols"] = [
                { wch: 8 },
                { wch: 15 },
                { wch: 15 },
                { wch: 20 },
                { wch: 15 },
                { wch: 12 },
                { wch: 10 },
                { wch: 30 },
                { wch: 20 },
                { wch: 15 },
                { wch: 12 },
                { wch: 50 },
                { wch: 12 }
            ];

            // Auto filter
            if (ws["!ref"]) {
                ws["!autofilter"] = {
                    ref: ws["!ref"]
                };
            }

            // Ép CCCD thành text
            if (ws["!ref"]) {
                const range =
                    XLSX.utils.decode_range(
                        ws["!ref"]
                    );

                for (
                    let r = 1;
                    r <= range.e.r;
                    r++
                ) {
                    const address =
                        XLSX.utils.encode_cell({
                            r,
                            c: 8
                        });

                    const cell =
                        ws[address];

                    if (cell) {
                        cell.t = "s";
                        cell.z = "@";
                    }
                }
            }

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

            const now =
                new Date();

            const yyyy =
                now.getFullYear();

            const mm =
                String(
                    now.getMonth() + 1
                ).padStart(2, "0");

            const dd =
                String(
                    now.getDate()
                ).padStart(2, "0");

            const hh =
                String(
                    now.getHours()
                ).padStart(2, "0");

            const mi =
                String(
                    now.getMinutes()
                ).padStart(2, "0");

            const fileName =
                `csdlksk_${yyyy}-${mm}-${dd}_${hh}-${mi}.xlsx`;

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
                `</span>`
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
                    `[KSK Export] API báo ${totalElementsExpected} hồ sơ (cả 2 nguồn), ` +
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
                `</span>`
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
