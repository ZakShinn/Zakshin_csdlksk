// ==UserScript==
// @name         Hỗ trợ xuất Excel DS CSDL KSK
// @namespace    https://hainghia.net/
// @version      1.9.6
// @description  Xuất toàn bộ danh sách khám sức khỏe từ csdlksk ra Excel
// @match        https://admin.csdlksk.vn/admin/operation/health-checkup*
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

    const API_URL =
        "https://api.emrhub.vn/api/checkup/health-examinations";

    const PAGE_SIZE = 100;

    // Số request chạy đồng thời.
    // Để 3-5 tương đối an toàn.
    const CONCURRENCY = 4;

    const MAX_RETRY = 3;

    const TARGET_PATH =
        "/admin/operation/health-checkup";

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
        retry = 0
    ) {
        const url =
            `${API_URL}` +
            `?page_number=${page}` +
            `&page_size=${PAGE_SIZE}`;

        try {
            const response =
                await gmRequest({
                    method: "GET",

                    url,

                    timeout: 30000,

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
                json =
                    JSON.parse(
                        response.responseText
                    );
            } catch {
                throw new Error(
                    "Response không phải JSON"
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

            return json;

        } catch (error) {
            if (retry < MAX_RETRY) {
                console.warn(
                    `[KSK Export] Trang ${page} lỗi. ` +
                    `Retry ${retry + 1}/${MAX_RETRY}`
                );

                await sleep(
                    800 * (retry + 1)
                );

                return fetchPage(
                    page,
                    token,
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
                            token
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
            // PAGE 1
            // ----------------------------------------------------

            setProgress(
                5,
                "📥 Đang đọc thông tin dữ liệu..."
            );

            const first =
                await fetchPage(
                    1,
                    token
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
                    "Response API không đúng cấu trúc."
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

            let allData =
                [...firstData.data];

            console.log(
                `[KSK Export] Tổng hồ sơ: ${totalElements}`
            );

            console.log(
                `[KSK Export] Tổng trang: ${totalPages}`
            );

            // ----------------------------------------------------
            // PAGE 2 -> END
            // ----------------------------------------------------

            if (totalPages > 1) {
                const pages =
                    [];

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
                        () => {
                            completed++;

                            const percent =
                                Math.round(
                                    (
                                        completed /
                                        totalPages
                                    ) *
                                    85
                                );

                            setProgress(
                                percent,
                                `📥 Đang tải dữ liệu...<br>` +
                                `${completed}/${totalPages} trang`
                            );
                        }
                    );

                const failed =
                    [];

                results
                    .sort(
                        (a, b) =>
                            a.page - b.page
                    )
                    .forEach(result => {
                        if (
                            result.success
                        ) {
                            const rows =
                                result
                                    .data
                                    ?.data
                                    ?.data;

                            if (
                                Array.isArray(
                                    rows
                                )
                            ) {
                                allData.push(
                                    ...rows
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
                        `Không tải được trang: ${failed.join(", ")}`
                    );
                }
            }

            // ----------------------------------------------------
            // LOẠI TRÙNG
            // ----------------------------------------------------

            setProgress(
                90,
                "🧹 Đang xử lý dữ liệu..."
            );

            const unique =
                new Map();

            for (const item of allData) {
                const key =
                    item.check_up_id ??
                    (
                        `${item.patient_id ?? ""}_` +
                        `${item.ngay_kham ?? ""}_` +
                        `${item.id_number ?? ""}`
                    );

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
                            item.detail_address ?? ""
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
                { wch: 50 }
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
                totalElements &&
                excelRows.length !==
                    totalElements
            ) {
                console.warn(
                    `[KSK Export] API báo ${totalElements} hồ sơ, ` +
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

    // Load ban đầu
    updateUI();

    // Theo dõi URL vì trang admin có thể là SPA
    let lastUrl =
        location.href;

    const observer =
        new MutationObserver(() => {
            if (
                location.href !==
                lastUrl
            ) {
                lastUrl =
                    location.href;

                setTimeout(
                    updateUI,
                    300
                );
            }

            // Nếu framework xóa nút thì tạo lại
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

})();
