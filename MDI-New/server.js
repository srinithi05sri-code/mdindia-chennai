require("dotenv").config();

const express = require("express");
const path = require("path");
const multer = require("multer");
const XLSX = require("xlsx");
const session = require("express-session");
const mysql = require("mysql2/promise");

const app = express();

const PORT = process.env.PORT || 3000;

// =====================================================
// MULTER - EXCEL UPLOAD
// =====================================================

const upload = multer({
    storage: multer.memoryStorage(),

    limits: {
        fileSize: 10 * 1024 * 1024
    },

    fileFilter: (req, file, cb) => {

        const allowed = [
            ".xlsx",
            ".xls"
        ];

        const ext =
            path.extname(file.originalname).toLowerCase();

        if (allowed.includes(ext)) {
            cb(null, true);
        } else {
            cb(
                new Error(
                    "Only Excel files (.xlsx, .xls) are allowed."
                )
            );
        }
    }
});

// =====================================================
// MYSQL CONNECTION
// =====================================================

console.log("MYSQLHOST:", process.env.MYSQLHOST);
console.log("MYSQLPORT:", process.env.MYSQLPORT);
console.log("MYSQLUSER:", process.env.MYSQLUSER);
console.log("MYSQLDATABASE:", process.env.MYSQLDATABASE);

const db = mysql.createPool({

    host: process.env.MYSQLHOST,
    port: Number(process.env.MYSQLPORT),
    user: process.env.MYSQLUSER,
    password: process.env.MYSQLPASSWORD,
    database: process.env.MYSQLDATABASE,

    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0
});

// =====================================================
// EXPRESS
// =====================================================

app.set(
    "view engine",
    "ejs"
);

app.set(
    "views",
    path.join(__dirname, "views")
);

app.use(
    express.urlencoded({
        extended: true
    })
);

app.use(express.json());

app.use(
    express.static(
        path.join(__dirname, "public")
    )
);

// =====================================================
// SESSION
// =====================================================

app.use(
    session({

        secret:
            process.env.SESSION_SECRET ||
            "mdi-claim-secret-key",

        resave: false,

        saveUninitialized: false,

        cookie: {
            maxAge: 1000 * 60 * 60,
            httpOnly: true
        }
    })
);

// =====================================================
// CONSTANTS
// =====================================================

const VALID_CLAIM_TYPES = [
    "IPD",
    "OPD",
    "Pre Post"
];

const VALID_STATUSES = [
    "Pending",
    "Approved",
    "Rejected",
    "Query",
    "Re-Query",
    "Query & Investigation",
    "Investigation",
    "Sent-Back",
    "Keep",
    "Other-Doctor/Executive",
    "ROD-Cancel"
];

// =====================================================
// ROLE NORMALIZER
// =====================================================

function normalizeRole(role) {

    return String(role || "")
        .trim()
        .toLowerCase();
}

// =====================================================
// STATUS NORMALIZER
// =====================================================

function normalizeStatus(status) {

    let value =
        String(status || "").trim();

    if (value === "") {
        return "Pending";
    }

    if (value === "SentBack") {
        value = "Sent-Back";
    }

    if (value === "ROD Cancel") {
        value = "ROD-Cancel";
    }

    if (value === "Investigation&Query") {
        value = "Query & Investigation";
    }

    if (value === "OtherDoctor/Executive") {
        value = "Other-Doctor/Executive";
    }

    return value;
}

// =====================================================
// CLAIM TYPE NORMALIZER
// =====================================================

function normalizeClaimType(value) {

    let claimType =
        String(value || "").trim();

    if (claimType === "") {
        return null;
    }

    const upper =
        claimType.toUpperCase();

    if (upper === "INPATIENT") {
        return "IPD";
    }

    if (upper === "OUTPATIENT") {
        return "OPD";
    }

    if (upper === "PREPOST") {
        return "Pre Post";
    }

    if (upper === "PRE POST") {
        return "Pre Post";
    }

    if (!VALID_CLAIM_TYPES.includes(claimType)) {

        throw new Error(
            `Invalid Claim Type '${claimType}'.`
        );
    }

    return claimType;
}

// =====================================================
// DATE VALIDATION
// =====================================================

function isValidDate(value) {

    if (!value) {
        return false;
    }

    return /^\d{4}-\d{2}-\d{2}$/.test(
        String(value)
    );
}

// =====================================================
// DATE RANGE VALIDATION
// =====================================================

function validateDateRange(fromDate, toDate) {

    if (!fromDate || !toDate) {

        return {
            valid: false,
            message:
                "From Date and To Date are required."
        };
    }

    if (
        !isValidDate(fromDate) ||
        !isValidDate(toDate)
    ) {

        return {
            valid: false,
            message:
                "Invalid date format."
        };
    }

    if (fromDate > toDate) {

        return {
            valid: false,
            message:
                "From Date cannot be greater than To Date."
        };
    }

    return {
        valid: true
    };
}

// =====================================================
// EXCEL DATE -> MYSQL DATE
// =====================================================

function excelDateToMysqlDate(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {
        return null;
    }

    // Excel serial date
    if (typeof value === "number") {

        const parsed =
            XLSX.SSF.parse_date_code(value);

        if (parsed) {

            const day =
                String(parsed.d).padStart(2, "0");

            const month =
                String(parsed.m).padStart(2, "0");

            const year =
                String(parsed.y);

            return `${year}-${month}-${day}`;
        }
    }

    // JS Date
    if (value instanceof Date) {

        if (!isNaN(value.getTime())) {

            const year =
                value.getFullYear();

            const month =
                String(
                    value.getMonth() + 1
                ).padStart(2, "0");

            const day =
                String(
                    value.getDate()
                ).padStart(2, "0");

            return `${year}-${month}-${day}`;
        }
    }

    const text =
        String(value).trim();

    if (!text) {
        return null;
    }

    // dd/mm/yyyy
    const indianDate =
        text.match(
            /^(\d{2})\/(\d{2})\/(\d{4})$/
        );

    if (indianDate) {

        return `${indianDate[3]}-${indianDate[2]}-${indianDate[1]}`;
    }

    // yyyy-mm-dd
    if (
        /^\d{4}-\d{2}-\d{2}$/.test(text)
    ) {
        return text;
    }

    return null;
}

// =====================================================
// MYSQL DATE -> DISPLAY DATE
// =====================================================

function formatDisplayDate(value) {

    if (!value) {
        return "-";
    }

    const text =
        String(value).trim();

    // yyyy-mm-dd
    if (
        /^\d{4}-\d{2}-\d{2}$/.test(text)
    ) {

        const [year, month, day] =
            text.split("-");

        return `${day}/${month}/${year}`;
    }

    return text;
}

// =====================================================
// EXCEL TIME -> MYSQL TIME
// =====================================================

function excelTimeToMysqlTime(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {
        return null;
    }

    // =================================================
    // Excel decimal time
    // Example:
    // 0.4083333333 = 09:48:00
    // =================================================

    if (typeof value === "number") {

        let totalSeconds =
            Math.round(value * 24 * 60 * 60);

        totalSeconds =
            totalSeconds % (24 * 60 * 60);

        if (totalSeconds < 0) {
            totalSeconds += 24 * 60 * 60;
        }

        const hour =
            Math.floor(
                totalSeconds / 3600
            );

        const minute =
            Math.floor(
                (totalSeconds % 3600) / 60
            );

        const second =
            totalSeconds % 60;

        return (
            String(hour).padStart(2, "0") +
            ":" +
            String(minute).padStart(2, "0") +
            ":" +
            String(second).padStart(2, "0")
        );
    }

    // =================================================
    // JS DATE
    // =================================================

    if (value instanceof Date) {

        if (!isNaN(value.getTime())) {

            return (
                String(
                    value.getHours()
                ).padStart(2, "0") +
                ":" +
                String(
                    value.getMinutes()
                ).padStart(2, "0") +
                ":" +
                String(
                    value.getSeconds()
                ).padStart(2, "0")
            );
        }
    }

    const text =
        String(value).trim();

    if (!text) {
        return null;
    }

    // HH:MM
    if (
        /^\d{2}:\d{2}$/.test(text)
    ) {

        return `${text}:00`;
    }

    // HH:MM:SS
    if (
        /^\d{2}:\d{2}:\d{2}$/.test(text)
    ) {

        return text;
    }

    // AM/PM
    const ampm =
        text.match(
            /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)$/i
        );

    if (ampm) {

        let hour =
            Number(ampm[1]);

        const minute =
            Number(ampm[2]);

        const second =
            Number(ampm[3] || 0);

        const period =
            ampm[4].toUpperCase();

        if (period === "AM" && hour === 12) {
            hour = 0;
        }

        if (period === "PM" && hour !== 12) {
            hour += 12;
        }

        return (
            String(hour).padStart(2, "0") +
            ":" +
            String(minute).padStart(2, "0") +
            ":" +
            String(second).padStart(2, "0")
        );
    }

    return text;
}

// =====================================================
// MYSQL TIME -> DISPLAY TIME
// =====================================================

function formatDisplayTime(value) {

    if (!value) {
        return "-";
    }

    let text =
        String(value).trim();

    // =================================================
    // If somehow decimal Excel time reached DB/display
    // =================================================

    if (
        !isNaN(Number(text)) &&
        Number(text) >= 0 &&
        Number(text) < 1
    ) {

        const mysqlTime =
            excelTimeToMysqlTime(
                Number(text)
            );

        text = mysqlTime;
    }

    // =================================================
    // HH:MM:SS
    // =================================================

    if (
        /^\d{2}:\d{2}:\d{2}$/.test(text)
    ) {

        const [
            hourRaw,
            minuteRaw,
            secondRaw
        ] = text.split(":");

        let hour =
            Number(hourRaw);

        const minute =
            Number(minuteRaw);

        const second =
            Number(secondRaw);

        const period =
            hour >= 12 ? "PM" : "AM";

        let displayHour =
            hour % 12;

        if (displayHour === 0) {
            displayHour = 12;
        }

        return (
            String(displayHour).padStart(2, "0") +
            ":" +
            String(minute).padStart(2, "0") +
            ":" +
            String(second).padStart(2, "0") +
            " " +
            period
        );
    }

    return text;
}

// =====================================================
// UPLOAD DATE/TIME FORMAT
// =====================================================

function formatUploadedAt(value) {

    if (!value) {
        return "-";
    }

    const date =
        new Date(value);

    if (isNaN(date.getTime())) {
        return String(value);
    }

    return date.toLocaleString(
        "en-IN",
        {
            day: "2-digit",
            month: "2-digit",
            year: "numeric",
            hour: "2-digit",
            minute: "2-digit",
            second: "2-digit",
            hour12: true,
            timeZone: "Asia/Kolkata"
        }
    );
}

// =====================================================
// DATABASE TEST
// =====================================================

async function testDatabase() {

    try {

        console.log(
            "========== MYSQL DEBUG =========="
        );

        console.log(
            "MYSQLHOST:",
            process.env.MYSQLHOST || "NOT SET"
        );

        console.log(
            "MYSQLPORT:",
            process.env.MYSQLPORT || "NOT SET"
        );

        console.log(
            "MYSQLUSER:",
            process.env.MYSQLUSER || "NOT SET"
        );

        console.log(
            "MYSQLDATABASE:",
            process.env.MYSQLDATABASE || "NOT SET"
        );

        console.log(
            "MYSQLPASSWORD:",
            process.env.MYSQLPASSWORD
                ? "SET"
                : "NOT SET"
        );

        console.log(
            "================================="
        );

        const connection =
            await db.getConnection();

        console.log(
            "MySQL Connected Successfully"
        );

        connection.release();

    } catch (error) {

        console.error(
            "MySQL Connection Failed"
        );

        console.error(
            "Error Code:",
            error.code
        );

        console.error(
            "Error Message:",
            error.message
        );
    }
}

// =====================================================
// LOGIN PAGE
// =====================================================

app.get("/", (req, res) => {

    res.render(
        "login",
        {
            error: null
        }
    );
});

// =====================================================
// LOGIN
// =====================================================

app.post("/login", async (req, res) => {

    const employee_id =
        String(
            req.body.employee_id || ""
        ).trim();

    const password =
        String(
            req.body.password || ""
        ).trim();

    if (!employee_id || !password) {

        return res.render("login", {
            error:
                "Employee ID and Password are required"
        });
    }

    try {

        const [users] =
            await db.query(
                `
                SELECT
                    id,
                    employee_id,
                    username,
                    password,
                    role,
                    department,
                    is_active
                FROM users
                WHERE LOWER(TRIM(employee_id))
                      = LOWER(TRIM(?))
                LIMIT 1
                `,
                [
                    employee_id
                ]
            );

        if (users.length === 0) {

            return res.render("login", {
                error:
                    "Invalid Employee ID or Password"
            });
        }

        const user =
            users[0];

        const dbPassword =
            String(
                user.password || ""
            ).trim();

        if (dbPassword !== password) {

            return res.render("login", {
                error:
                    "Invalid Employee ID or Password"
            });
        }

        const activeValue =
            String(
                user.is_active
            )
                .trim()
                .toLowerCase();

        const isActive =
            activeValue === "1" ||
            activeValue === "true";

        if (!isActive) {

            return res.render("login", {
                error:
                    "Your account is inactive"
            });
        }

        const role =
            normalizeRole(
                user.role
            );

        req.session.user = {

            id:
                user.id,

            employee_id:
                user.employee_id,

            username:
                user.username,

            role:
                role,

            department:
                user.department
        };

        if (role === "admin") {
            return res.redirect("/admin");
        }

        if (role === "upload") {
            return res.redirect("/upload");
        }

        if (role === "user") {
            return res.redirect("/user");
        }

        return res.render("login", {
            error:
                "Invalid user role"
        });

    } catch (error) {

        console.error(
            "LOGIN DATABASE ERROR:",
            error
        );

        return res.render("login", {
            error:
                "Server error. Please try again."
        });
    }
});

// =====================================================
// UPLOAD DASHBOARD
// =====================================================

app.get(
    "/upload",
    async (req, res) => {

        if (
            !req.session.user ||
            normalizeRole(
                req.session.user.role
            ) !== "upload"
        ) {
            return res.redirect("/");
        }

        try {

            const [uploads] =
                await db.query(
                    `
                    SELECT

                        ub.id,

                        ub.file_name,

                        ub.uploaded_at,

                        ub.total_claims,

                        CASE
                            WHEN ub.status = 'ACTIVE'
                            THEN 'ACTIVE'
                            ELSE 'DELETED'
                        END AS status

                    FROM upload_batches ub

                    ORDER BY ub.id DESC
                    `
                );

            const formattedUploads =
                uploads.map(uploadRow => ({

                    ...uploadRow,

                    formatted_uploaded_at:
                        formatUploadedAt(
                            uploadRow.uploaded_at
                        )
                }));

            return res.render(
                "upload-dashboard",
                {

                    user:
                        req.session.user,

                    uploads:
                        formattedUploads
                }
            );

        } catch (error) {

            console.error(
                "UPLOAD DASHBOARD ERROR:",
                error
            );

            return res.status(500).send(`
                <h2>Database Error</h2>
                <pre>${error.message}</pre>
            `);
        }
    }
);

// =====================================================
// UPLOAD EXCEL
// =====================================================

app.post(
    "/upload-excel",
    upload.single("excelFile"),
    async (req, res) => {

        if (
            !req.session.user ||
            normalizeRole(
                req.session.user.role
            ) !== "upload"
        ) {
            return res.redirect("/");
        }

        if (!req.file) {

            return res.status(400).send(
                "Please select an Excel file."
            );
        }

        let connection;

        try {

            // =================================================
            // READ EXCEL
            // =================================================

            const workbook =
                XLSX.read(
                    req.file.buffer,
                    {
                        type: "buffer",
                        cellDates: true
                    }
                );

            const sheetName =
                workbook.SheetNames[0];

            const sheet =
                workbook.Sheets[sheetName];

            const rows =
                XLSX.utils.sheet_to_json(
                    sheet,
                    {
                        defval: ""
                    }
                );

            if (rows.length === 0) {

                return res.status(400).send(`
                    <h2>Excel file is empty</h2>
                    <a href="/upload">
                        Back to Upload
                    </a>
                `);
            }

            // =================================================
            // REQUIRED COLUMNS
            // =================================================

            const requiredColumns = [

                "CLAIM_REF_NO",
                "INWARD_NO",
                "POLICY_NO",
                "Vertical",
                "Additional Deduction",
                "CLAIM_AMT",
                "AL_AMT",
                "CLAIM_CLASS",
                "Hospital Code",
                "Type of MOU",
                "Diagnosis",
                "Diagnosis 2",
                "POLICY_NAME",
                "Queue",
                "Ageing",
                "Date",
                "Time",
                "Today Status",
                "I3 Status",
                "Full qc",
                "RELATION",
                "HNF",
                "User ID",
                "User Name",
                "Claim Type",
                "ILOM ID",
                "Approve AMT",
                "Status",
                "Remark",
                "Deduction AMT",
                "inter. Doc & Exe",
                "lot",
                "platform"
            ];

            const excelColumns =
                Object.keys(rows[0]);

            const missingColumns =
                requiredColumns.filter(
                    column =>
                        !excelColumns.includes(
                            column
                        )
                );

            if (missingColumns.length > 0) {

                return res.status(400).send(`
                    <h2>
                        Invalid Excel Format
                    </h2>

                    <p>
                        Missing columns:
                    </p>

                    <ul>
                        ${missingColumns
                            .map(
                                column =>
                                    `<li>${column}</li>`
                            )
                            .join("")}
                    </ul>

                    <a href="/upload">
                        Back to Upload
                    </a>
                `);
            }

            // =================================================
            // DB CONNECTION
            // =================================================

            connection =
                await db.getConnection();

            await connection.beginTransaction();

            // =================================================
            // CREATE BATCH
            // =================================================

            const [batchResult] =
                await connection.query(
                    `
                    INSERT INTO upload_batches
                    (
                        file_name,
                        uploaded_by,
                        total_claims,
                        status
                    )
                    VALUES
                    (
                        ?,
                        ?,
                        ?,
                        'ACTIVE'
                    )
                    `,
                    [
                        req.file.originalname,
                        req.session.user.id,
                        rows.length
                    ]
                );

            const batchId =
                batchResult.insertId;

            // =================================================
            // INSERT CLAIMS
            // =================================================

            for (const row of rows) {

                const claimRefNo =
                    String(
                        row["CLAIM_REF_NO"] || ""
                    ).trim();

                const inwardNo =
                    String(
                        row["INWARD_NO"] || ""
                    ).trim();

                const policyNo =
                    String(
                        row["POLICY_NO"] || ""
                    ).trim();

                const claimAmount =
                    parseFloat(
                        row["CLAIM_AMT"]
                    ) || 0;

                const vertical =
                    String(
                        row["Vertical"] || ""
                    ).trim();

                const department =
                    vertical;

                // =================================================
                // USER ID
                // =================================================

                const employeeId =
                    String(
                        row["User ID"] || ""
                    ).trim();

                let assignedUserId = null;

                let userName = null;

                if (employeeId) {

                    const [userRows] =
                        await connection.query(
                            `
                            SELECT
                                employee_id,
                                username
                            FROM users
                            WHERE employee_id = ?
                            AND LOWER(TRIM(role)) = 'user'
                            AND is_active = TRUE
                            LIMIT 1
                            `,
                            [
                                employeeId
                            ]
                        );

                    if (userRows.length === 0) {

                        throw new Error(
                            `Employee ID '${employeeId}' not found or inactive.`
                        );
                    }

                    assignedUserId =
                        userRows[0].employee_id;

                    userName =
                        userRows[0].username;
                }

                const additionalDeduction =
                    parseFloat(
                        row["Additional Deduction"]
                    ) || 0;

                const alAmount =
                    parseFloat(
                        row["AL_AMT"]
                    ) || 0;

                const claimClass =
                    String(
                        row["CLAIM_CLASS"] || ""
                    ).trim();

                const hospitalCode =
                    String(
                        row["Hospital Code"] || ""
                    ).trim();

                const typeOfMou =
                    String(
                        row["Type of MOU"] || ""
                    ).trim();

                const diagnosis =
                    String(
                        row["Diagnosis"] || ""
                    ).trim();

                const diagnosis2 =
                    String(
                        row["Diagnosis 2"] || ""
                    ).trim();

                const policyName =
                    String(
                        row["POLICY_NAME"] || ""
                    ).trim();

                const queue =
                    String(
                        row["Queue"] || ""
                    ).trim();

                const ageing =
                    String(
                        row["Ageing"] || ""
                    ).trim();

                // =================================================
                // IMPORTANT:
                // Store DATE as YYYY-MM-DD
                // Store TIME as HH:MM:SS
                // =================================================

                const claimDate =
                    excelDateToMysqlDate(
                        row["Date"]
                    );

                const claimTime =
                    excelTimeToMysqlTime(
                        row["Time"]
                    );

                const todayStatus =
                    String(
                        row["Today Status"] || ""
                    ).trim();

                const i3Status =
                    String(
                        row["I3 Status"] || ""
                    ).trim();

                const fullQc =
                    String(
                        row["Full qc"] || ""
                    ).trim();

                const relation =
                    String(
                        row["RELATION"] || ""
                    ).trim();

                const hnf =
                    String(
                        row["HNF"] || ""
                    ).trim();

                const ilomId =
                    String(
                        row["ILOM ID"] || ""
                    ).trim();

                const approveAmount =
                    parseFloat(
                        row["Approve AMT"]
                    ) || 0;

                const remark =
                    String(
                        row["Remark"] || ""
                    ).trim();

                const deductionAmount =
                    parseFloat(
                        row["Deduction AMT"]
                    ) || 0;

                const interDocExe =
                    String(
                        row["inter. Doc & Exe"] || ""
                    ).trim();

                const lot =
                    String(
                        row["lot"] ||
                        row["LOT"] ||
                        ""
                    ).trim();

                const platform =
                    String(
                        row["platform"] ||
                        row["Platform"] ||
                        ""
                    ).trim();

                // =================================================
                // CLAIM TYPE
                // =================================================

                const claimType =
                    normalizeClaimType(
                        row["Claim Type"]
                    );

                // =================================================
                // STATUS
                // =================================================

                const claimStatus =
                    normalizeStatus(
                        row["Status"]
                    );

                if (
                    !VALID_STATUSES.includes(
                        claimStatus
                    )
                ) {

                    throw new Error(
                        `Invalid Status '${claimStatus}' for Claim ${claimRefNo}.`
                    );
                }

                // =================================================
                // INSERT CLAIM
                //
                // IMPORTANT:
                // uploaded_at REMOVED
                // It belongs to upload_batches.
                // =================================================

                await connection.query(
                    `
                    INSERT INTO claims
                    (
                        upload_batch_id,
                        claim_ref_no,
                        inward_no,
                        policy_no,
                        claim_amount,
                        vertical,
                        department,
                        assigned_user_id,
                        user_name,
                        claim_type,
                        claim_status,
                        user_remark,
                        additional_deduction,
                        al_amount,
                        claim_class,
                        hospital_code,
                        type_of_mou,
                        diagnosis,
                        diagnosis_2,
                        policy_name,
                        queue,
                        ageing,
                        claim_date,
                        claim_time,
                        today_status,
                        i3_status,
                        full_qc,
                        relation,
                        hnf,
                        ilom_id,
                        approve_amount,
                        deduction_amount,
                        inter_doc_exe,
                        lot,
                        platform
                    )
                    VALUES
                    (
                        ?, ?, ?, ?, ?, ?,
                        ?, ?, ?, ?, ?, ?,
                        ?, ?, ?, ?, ?, ?,
                        ?, ?, ?, ?, ?, ?,
                        ?, ?, ?, ?, ?, ?,
                        ?, ?, ?, ?, ?
                    )
                    `,
                    [
                        batchId,
                        claimRefNo,
                        inwardNo,
                        policyNo,
                        claimAmount,
                        vertical,
                        department,
                        assignedUserId,
                        userName,
                        claimType,
                        claimStatus,
                        remark,
                        additionalDeduction,
                        alAmount,
                        claimClass,
                        hospitalCode,
                        typeOfMou,
                        diagnosis,
                        diagnosis2,
                        policyName,
                        queue,
                        ageing,
                        claimDate,
                        claimTime,
                        todayStatus,
                        i3Status,
                        fullQc,
                        relation,
                        hnf,
                        ilomId,
                        approveAmount,
                        deductionAmount,
                        interDocExe,
                        lot,
                        platform
                    ]
                );
            }

            await connection.commit();

            console.log(
                "EXCEL UPLOAD SUCCESS"
            );

            return res.redirect("/upload");

        } catch (error) {

            if (connection) {

                try {
                    await connection.rollback();
                } catch (rollbackError) {
                    console.error(
                        "ROLLBACK ERROR:",
                        rollbackError
                    );
                }
            }

            console.error(
                "EXCEL UPLOAD ERROR:",
                error
            );

            return res.status(500).send(`
                <h2>
                    Excel Upload Failed
                </h2>

                <pre>
${error.message}
                </pre>

                <br>

                <a href="/upload">
                    Back to Upload
                </a>
            `);

        } finally {

            if (connection) {
                connection.release();
            }
        }
    }
);

// =====================================================
// DELETE UPLOAD
// =====================================================

app.post(
    "/delete-upload/:id",
    async (req, res) => {

        if (
            !req.session.user ||
            normalizeRole(
                req.session.user.role
            ) !== "upload"
        ) {
            return res.redirect("/");
        }

        const batchId =
            req.params.id;

        let connection;

        try {

            connection =
                await db.getConnection();

            await connection.beginTransaction();

            await connection.query(
                `
                DELETE FROM claims
                WHERE upload_batch_id = ?
                `,
                [
                    batchId
                ]
            );

            await connection.query(
                `
                UPDATE upload_batches
                SET status = 'DELETED'
                WHERE id = ?
                AND status = 'ACTIVE'
                `,
                [
                    batchId
                ]
            );

            await connection.commit();

            return res.redirect("/upload");

        } catch (error) {

            if (connection) {
                await connection.rollback();
            }

            console.error(
                "DELETE UPLOAD ERROR:",
                error
            );

            return res.status(500).send(`
                <h2>
                    Delete Failed
                </h2>

                <pre>
${error.message}
                </pre>

                <br>

                <a href="/upload">
                    Back to Upload
                </a>
            `);

        } finally {

            if (connection) {
                connection.release();
            }
        }
    }
);

// =====================================================
// ADMIN DASHBOARD
// =====================================================

app.get(
    "/admin",
    async (req, res) => {

        if (
            !req.session.user ||
            normalizeRole(
                req.session.user.role
            ) !== "admin"
        ) {
            return res.redirect("/");
        }

        try {

            // =================================================
            // OVERALL CLAIM COUNTS
            // =================================================

            const [[summary]] =
                await db.query(`
                    SELECT

                        COUNT(*) AS total,

                        SUM(
                            claim_status = 'Pending'
                        ) AS pending,

                        SUM(
                            claim_status = 'Approved'
                        ) AS approved,

                        SUM(
                            claim_status = 'Rejected'
                        ) AS rejected,

                        SUM(
                            claim_status = 'Query'
                        ) AS query_count,

                        SUM(
                            claim_status = 'Re-Query'
                        ) AS requery,

                        SUM(
                            claim_status =
                            'Query & Investigation'
                        ) AS investigation_query,

                        SUM(
                            claim_status =
                            'Investigation'
                        ) AS investigation,

                        SUM(
                            claim_status =
                            'Sent-Back'
                        ) AS sent_back,

                        SUM(
                            claim_status = 'Keep'
                        ) AS keep_count,

                        SUM(
                            claim_status =
                            'Other-Doctor/Executive'
                        ) AS other_doctor_executive,

                        SUM(
                            claim_status =
                            'ROD-Cancel'
                        ) AS rod_cancel,

                        SUM(
                            claim_status <> 'Pending'
                        ) AS processed

                    FROM claims
                `);

            // =================================================
            // TOTAL ACTIVE USERS
            // =================================================

            const [[userCount]] =
                await db.query(`
                    SELECT
                        COUNT(*) AS count
                    FROM users
                    WHERE LOWER(TRIM(role)) = 'user'
                    AND is_active = TRUE
                `);

            // =================================================
            // USER LIST
            // =================================================

            const [userList] =
                await db.query(`
                    SELECT
                        id,
                        employee_id,
                        username,
                        department,
                        is_active
                    FROM users
                    WHERE LOWER(TRIM(role)) = 'user'
                    ORDER BY username
                `);

            // =================================================
            // PROCESS SUMMARY
            // =================================================

            const [processSummary] =
                await db.query(`
                    SELECT

                        COALESCE(
                            c.platform,
                            '-'
                        ) AS platform,

                        COALESCE(
                            u.employee_id,
                            c.assigned_user_id,
                            '-'
                        ) AS employee_id,

                        COALESCE(
                            u.username,
                            c.user_name,
                            '-'
                        ) AS user_name,

                        COUNT(*) AS total_allocated,

                        SUM(
                            c.claim_status = 'Pending'
                        ) AS pending,

                        SUM(
                            c.claim_status = 'Approved'
                        ) AS approved,

                        SUM(
                            c.claim_status = 'Rejected'
                        ) AS rejected,

                        SUM(
                            c.claim_status = 'Query'
                        ) AS query_count,

                        SUM(
                            c.claim_status = 'Re-Query'
                        ) AS requery,

                        SUM(
                            c.claim_status =
                            'Query & Investigation'
                        ) AS investigation_query,

                        SUM(
                            c.claim_status =
                            'Investigation'
                        ) AS investigation,

                        SUM(
                            c.claim_status =
                            'Sent-Back'
                        ) AS sent_back,

                        SUM(
                            c.claim_status = 'Keep'
                        ) AS keep_count,

                        SUM(
                            c.claim_status =
                            'Other-Doctor/Executive'
                        ) AS other_doctor_executive,

                        SUM(
                            c.claim_status =
                            'ROD-Cancel'
                        ) AS rod_cancel,

                        SUM(
                            c.claim_status <> 'Pending'
                        ) AS total_productivity

                    FROM claims c

                    LEFT JOIN users u
                        ON TRIM(
                            c.assigned_user_id
                        )
                        =
                        TRIM(
                            u.employee_id
                        )

                    GROUP BY
                        c.platform,
                        COALESCE(
                            u.employee_id,
                            c.assigned_user_id,
                            '-'
                        ),
                        COALESCE(
                            u.username,
                            c.user_name,
                            '-'
                        )

                    ORDER BY
                        c.platform,
                        user_name
                `);

            const adminUser = {
                ...req.session.user,
                name:
                    req.session.user.username
            };

            return res.render(
                "admin-dashboard",
                {

                    user:
                        adminUser,

                    counts: {

                        total:
                            Number(
                                summary.total || 0
                            ),

                        pending:
                            Number(
                                summary.pending || 0
                            ),

                        approved:
                            Number(
                                summary.approved || 0
                            ),

                        rejected:
                            Number(
                                summary.rejected || 0
                            ),

                        query:
                            Number(
                                summary.query_count || 0
                            ),

                        requery:
                            Number(
                                summary.requery || 0
                            ),

                        investigationQuery:
                            Number(
                                summary.investigation_query || 0
                            ),

                        investigation:
                            Number(
                                summary.investigation || 0
                            ),

                        sentBack:
                            Number(
                                summary.sent_back || 0
                            ),

                        keep:
                            Number(
                                summary.keep_count || 0
                            ),

                        otherDoctorExecutive:
                            Number(
                                summary.other_doctor_executive || 0
                            ),

                        rodCancel:
                            Number(
                                summary.rod_cancel || 0
                            ),

                        processed:
                            Number(
                                summary.processed || 0
                            ),

                        users:
                            Number(
                                userCount.count || 0
                            )
                    },

                    userList:
                        userList,

                    processSummary:
                        processSummary
                }
            );

        } catch (error) {

            console.error(
                "ADMIN DASHBOARD ERROR:",
                error
            );

            return res.status(500).send(`
                <h2>Admin Dashboard Error</h2>
                <pre>${error.message}</pre>
                <br>
                <a href="/admin">
                    Back to Admin
                </a>
            `);
        }
    }
);

// =====================================================
// ADMIN DOWNLOAD CLAIMS
// =====================================================
//
// IMPORTANT:
// SELECT * REMOVED
//
// So these columns will NOT come:
// id
// upload_batch_id
// created_at
// updated_at
//
// =====================================================

app.get(
    "/admin/download-claims",
    async (req, res) => {

        if (
            !req.session.user ||
            normalizeRole(
                req.session.user.role
            ) !== "admin"
        ) {
            return res.redirect("/");
        }

        try {

            const [claims] =
                await db.query(`
                    SELECT

                        claim_ref_no AS CLAIM_REF_NO,

                        inward_no AS INWARD_NO,

                        policy_no AS POLICY_NO,

                        vertical AS Vertical,

                        additional_deduction
                            AS 'Additional Deduction',

                        claim_amount AS CLAIM_AMT,

                        al_amount AS AL_AMT,

                        claim_class AS CLAIM_CLASS,

                        hospital_code AS 'Hospital Code',

                        type_of_mou AS 'Type of MOU',

                        diagnosis AS Diagnosis,

                        diagnosis_2 AS 'Diagnosis 2',

                        policy_name AS POLICY_NAME,

                        queue AS Queue,

                        ageing AS Ageing,

                        claim_date AS claim_date,

                        claim_time AS claim_time,

                        today_status AS 'Today Status',

                        i3_status AS 'I3 Status',

                        full_qc AS 'Full qc',

                        relation AS RELATION,

                        hnf AS HNF,

                        assigned_user_id AS 'User ID',

                        user_name AS 'User Name',

                        claim_type AS 'Claim Type',

                        ilom_id AS 'ILOM ID',

                        approve_amount AS 'Approve AMT',

                        claim_status AS Status,

                        user_remark AS Remark,

                        deduction_amount
                            AS 'Deduction AMT',

                        inter_doc_exe
                            AS 'inter. Doc & Exe',

                        lot AS lot,

                        platform AS platform

                    FROM claims

                    ORDER BY id
                `);

            // =================================================
            // FORMAT FOR EXCEL
            // =================================================

            const excelData =
                claims.map(row => ({

                    "CLAIM_REF_NO":
                        row.CLAIM_REF_NO || "",

                    "INWARD_NO":
                        row.INWARD_NO || "",

                    "POLICY_NO":
                        row.POLICY_NO || "",

                    "Vertical":
                        row.Vertical || "",

                    "Additional Deduction":
                        Number(
                            row["Additional Deduction"] || 0
                        ),

                    "CLAIM_AMT":
                        Number(
                            row.CLAIM_AMT || 0
                        ),

                    "AL_AMT":
                        Number(
                            row.AL_AMT || 0
                        ),

                    "CLAIM_CLASS":
                        row.CLAIM_CLASS || "",

                    "Hospital Code":
                        row["Hospital Code"] || "",

                    "Type of MOU":
                        row["Type of MOU"] || "",

                    "Diagnosis":
                        row.Diagnosis || "",

                    "Diagnosis 2":
                        row["Diagnosis 2"] || "",

                    "POLICY_NAME":
                        row.POLICY_NAME || "",

                    "Queue":
                        row.Queue || "",

                    "Ageing":
                        row.Ageing || "",

                    "Date":
                        formatDisplayDate(
                            row.claim_date
                        ),

                    "Time":
                        formatDisplayTime(
                            row.claim_time
                        ),

                    "Today Status":
                        row["Today Status"] || "",

                    "I3 Status":
                        row["I3 Status"] || "",

                    "Full qc":
                        row["Full qc"] || "",

                    "RELATION":
                        row.RELATION || "",

                    "HNF":
                        row.HNF || "",

                    "User ID":
                        row["User ID"] || "",

                    "User Name":
                        row["User Name"] || "",

                    "Claim Type":
                        row["Claim Type"] || "",

                    "ILOM ID":
                        row["ILOM ID"] || "",

                    "Approve AMT":
                        Number(
                            row["Approve AMT"] || 0
                        ),

                    "Status":
                        row.Status || "Pending",

                    "Remark":
                        row.Remark || "",

                    "Deduction AMT":
                        Number(
                            row["Deduction AMT"] || 0
                        ),

                    "inter. Doc & Exe":
                        row["inter. Doc & Exe"] || "",

                    "lot":
                        row.lot || "",

                    "platform":
                        row.platform || ""
                }));

            const worksheet =
                XLSX.utils.json_to_sheet(
                    excelData
                );

            worksheet["!cols"] = [

                { wch: 20 },
                { wch: 18 },
                { wch: 18 },
                { wch: 15 },
                { wch: 20 },
                { wch: 15 },
                { wch: 15 },
                { wch: 18 },
                { wch: 18 },
                { wch: 18 },
                { wch: 25 },
                { wch: 25 },
                { wch: 25 },
                { wch: 15 },
                { wch: 12 },
                { wch: 15 },
                { wch: 15 },
                { wch: 20 },
                { wch: 18 },
                { wch: 15 },
                { wch: 15 },
                { wch: 15 },
                { wch: 18 },
                { wch: 25 },
                { wch: 15 },
                { wch: 18 },
                { wch: 18 },
                { wch: 20 },
                { wch: 30 },
                { wch: 18 },
                { wch: 15 },
                { wch: 15 },
                { wch: 15 }
            ];

            const workbook =
                XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(
                workbook,
                worksheet,
                "Claims"
            );

            const buffer =
                XLSX.write(
                    workbook,
                    {
                        type: "buffer",
                        bookType: "xlsx"
                    }
                );

            res.setHeader(
                "Content-Disposition",
                "attachment; filename=updated_claims.xlsx"
            );

            res.setHeader(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            );

            return res.send(buffer);

        } catch (error) {

            console.error(
                "CLAIM DOWNLOAD ERROR:",
                error
            );

            return res.status(500).send(
                "Failed to download claims"
            );
        }
    }
);

// =====================================================
// DOWNLOAD PROCESS SUMMARY
// =====================================================

app.get(
    "/admin/download-process-summary",
    async (req, res) => {

        if (
            !req.session.user ||
            normalizeRole(
                req.session.user.role
            ) !== "admin"
        ) {
            return res.redirect("/");
        }

        try {

            const [rows] =
                await db.query(`
                    SELECT

                        COALESCE(
                            c.platform,
                            '-'
                        ) AS platform,

                        COALESCE(
                            u.employee_id,
                            c.assigned_user_id,
                            '-'
                        ) AS employee_id,

                        COALESCE(
                            u.username,
                            c.user_name,
                            '-'
                        ) AS user_name,

                        COUNT(*) AS total_allocated,

                        SUM(
                            c.claim_status = 'Approved'
                        ) AS approved,

                        SUM(
                            c.claim_status = 'Rejected'
                        ) AS rejected,

                        SUM(
                            c.claim_status = 'Query'
                        ) AS query_count,

                        SUM(
                            c.claim_status = 'Re-Query'
                        ) AS requery,

                        SUM(
                            c.claim_status =
                            'Query & Investigation'
                        ) AS investigation_query,

                        SUM(
                            c.claim_status =
                            'Investigation'
                        ) AS investigation,

                        SUM(
                            c.claim_status =
                            'Sent-Back'
                        ) AS sent_back,

                        SUM(
                            c.claim_status = 'Keep'
                        ) AS keep_count,

                        SUM(
                            c.claim_status =
                            'Other-Doctor/Executive'
                        ) AS other_doctor_executive,

                        SUM(
                            c.claim_status =
                            'ROD-Cancel'
                        ) AS rod_cancel,

                        SUM(
                            c.claim_status <> 'Pending'
                        ) AS total_productivity,

                        SUM(
                            c.claim_status = 'Pending'
                        ) AS pending

                    FROM claims c

                    LEFT JOIN users u
                        ON TRIM(
                            c.assigned_user_id
                        )
                        =
                        TRIM(
                            u.employee_id
                        )

                    GROUP BY
                        c.platform,
                        COALESCE(
                            u.employee_id,
                            c.assigned_user_id,
                            '-'
                        ),
                        COALESCE(
                            u.username,
                            c.user_name,
                            '-'
                        )

                    ORDER BY
                        c.platform,
                        user_name
                `);

            const workbook =
                XLSX.utils.book_new();

            const excelData =
                rows.map(row => ({

                    "Platform":
                        row.platform || "-",

                    "Employee ID":
                        row.employee_id || "-",

                    "User Name":
                        row.user_name || "-",

                    "Total Allocated":
                        Number(
                            row.total_allocated || 0
                        ),

                    "Approved":
                        Number(
                            row.approved || 0
                        ),

                    "Rejected":
                        Number(
                            row.rejected || 0
                        ),

                    "Query":
                        Number(
                            row.query_count || 0
                        ),

                    "Re-Query":
                        Number(
                            row.requery || 0
                        ),

                    "Query + Investigation":
                        Number(
                            row.investigation_query || 0
                        ),

                    "Total Productivity":
                        Number(
                            row.total_productivity || 0
                        ),

                    "Investigation":
                        Number(
                            row.investigation || 0
                        ),

                    "Sent Back":
                        Number(
                            row.sent_back || 0
                        ),

                    "Keep":
                        Number(
                            row.keep_count || 0
                        ),

                    "Other Doctor & Executive":
                        Number(
                            row.other_doctor_executive || 0
                        ),

                    "ROD Cancel":
                        Number(
                            row.rod_cancel || 0
                        ),

                    "Pending":
                        Number(
                            row.pending || 0
                        )
                }));

            const worksheet =
                XLSX.utils.json_to_sheet(
                    excelData
                );

            worksheet["!cols"] = [

                { wch: 18 },
                { wch: 18 },
                { wch: 25 },
                { wch: 18 },
                { wch: 12 },
                { wch: 12 },
                { wch: 12 },
                { wch: 12 },
                { wch: 22 },
                { wch: 20 },
                { wch: 18 },
                { wch: 15 },
                { wch: 12 },
                { wch: 25 },
                { wch: 15 },
                { wch: 12 }
            ];

            XLSX.utils.book_append_sheet(
                workbook,
                worksheet,
                "Process Summary"
            );

            const buffer =
                XLSX.write(
                    workbook,
                    {
                        type: "buffer",
                        bookType: "xlsx"
                    }
                );

            res.setHeader(
                "Content-Disposition",
                "attachment; filename=process-summary.xlsx"
            );

            res.setHeader(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            );

            return res.send(buffer);

        } catch (error) {

            console.error(
                "PROCESS SUMMARY DOWNLOAD ERROR:",
                error
            );

            return res.status(500).send(`
                <h2>
                    Process Summary Download Error
                </h2>

                <pre>
${error.message}
                </pre>

                <br>

                <a href="/admin">
                    Back to Admin
                </a>
            `);
        }
    }
);

// =====================================================
// CREATE USER
// =====================================================

app.post(
    "/admin/create-user",
    async (req, res) => {

        if (
            !req.session.user ||
            normalizeRole(
                req.session.user.role
            ) !== "admin"
        ) {
            return res.redirect("/");
        }

        const employeeId =
            String(
                req.body.employeeId || ""
            ).trim();

        const username =
            String(
                req.body.username || ""
            ).trim();

        const password =
            String(
                req.body.password || ""
            ).trim();

        const department =
            String(
                req.body.department || ""
            ).trim();

        if (
            !employeeId ||
            !username ||
            !password
        ) {

            return res.status(400).send(`
                <h2>Create User Failed</h2>

                <p>
                    Employee ID, Username and Password
                    are required.
                </p>

                <a href="/admin">
                    Back to Admin
                </a>
            `);
        }

        try {

            const [employeeExists] =
                await db.query(
                    `
                    SELECT id
                    FROM users
                    WHERE LOWER(TRIM(employee_id))
                        =
                        LOWER(TRIM(?))
                    LIMIT 1
                    `,
                    [
                        employeeId
                    ]
                );

            if (employeeExists.length > 0) {

                return res.status(400).send(`
                    <h2>Create User Failed</h2>

                    <p>
                        Employee ID
                        <b>${employeeId}</b>
                        already exists.
                    </p>

                    <a href="/admin">
                        Back to Admin
                    </a>
                `);
            }

            const [usernameExists] =
                await db.query(
                    `
                    SELECT id
                    FROM users
                    WHERE LOWER(TRIM(username))
                        =
                        LOWER(TRIM(?))
                    LIMIT 1
                    `,
                    [
                        username
                    ]
                );

            if (usernameExists.length > 0) {

                return res.status(400).send(`
                    <h2>Create User Failed</h2>

                    <p>
                        Username
                        <b>${username}</b>
                        already exists.
                    </p>

                    <a href="/admin">
                        Back to Admin
                    </a>
                `);
            }

            await db.query(
                `
                INSERT INTO users
                (
                    employee_id,
                    username,
                    password,
                    role,
                    department,
                    is_active
                )
                VALUES
                (
                    ?,
                    ?,
                    ?,
                    'user',
                    ?,
                    TRUE
                )
                `,
                [
                    employeeId,
                    username,
                    password,
                    department || null
                ]
            );

            return res.redirect("/admin");

        } catch (error) {

            console.error(
                "CREATE USER ERROR:",
                error
            );

            return res.status(500).send(`
                <h2>Create User Failed</h2>

                <pre>
${error.message}
                </pre>

                <br>

                <a href="/admin">
                    Back to Admin
                </a>
            `);
        }
    }
);

// =====================================================
// REASSIGN CLAIMS
// =====================================================

app.post(
    "/admin/reassign",
    async (req, res) => {

        if (
            !req.session.user ||
            normalizeRole(
                req.session.user.role
            ) !== "admin"
        ) {
            return res.redirect("/");
        }

        const oldUserId =
            String(
                req.body.oldUserId || ""
            ).trim();

        const newUserId =
            String(
                req.body.newUserId || ""
            ).trim();

        if (!oldUserId || !newUserId) {

            return res.status(400).send(`
                <h2>Invalid User Selection</h2>

                <a href="/admin">
                    Back to Admin
                </a>
            `);
        }

        if (oldUserId === newUserId) {

            return res.status(400).send(`
                <h2>Invalid Reassignment</h2>

                <p>
                    Old User and New User cannot be same.
                </p>

                <a href="/admin">
                    Back to Admin
                </a>
            `);
        }

        let connection;

        try {

            connection =
                await db.getConnection();

            await connection.beginTransaction();

            const [oldUser] =
                await connection.query(
                    `
                    SELECT employee_id
                    FROM users
                    WHERE id = ?
                    AND LOWER(TRIM(role)) = 'user'
                    LIMIT 1
                    `,
                    [
                        oldUserId
                    ]
                );

            if (oldUser.length === 0) {

                throw new Error(
                    "Old user not found."
                );
            }

            const oldEmployeeId =
                String(
                    oldUser[0].employee_id
                ).trim();

            const [newUser] =
                await connection.query(
                    `
                    SELECT
                        employee_id,
                        username
                    FROM users
                    WHERE id = ?
                    AND LOWER(TRIM(role)) = 'user'
                    AND is_active = TRUE
                    LIMIT 1
                    `,
                    [
                        newUserId
                    ]
                );

            if (newUser.length === 0) {

                throw new Error(
                    "New user not found or inactive."
                );
            }

            const newEmployeeId =
                String(
                    newUser[0].employee_id
                ).trim();

            const newUsername =
                newUser[0].username;

            await connection.query(
                `
                UPDATE claims
                SET
                    assigned_user_id = ?,
                    user_name = ?,
                    updated_at = NOW()
                WHERE
                    TRIM(assigned_user_id)
                    =
                    TRIM(?)
                `,
                [
                    newEmployeeId,
                    newUsername,
                    oldEmployeeId
                ]
            );

            await connection.commit();

            return res.redirect("/admin");

        } catch (error) {

            if (connection) {
                await connection.rollback();
            }

            console.error(
                "REASSIGN ERROR:",
                error
            );

            return res.status(500).send(`
                <h2>Reassignment Failed</h2>

                <pre>
${error.message}
                </pre>

                <br>

                <a href="/admin">
                    Back to Admin
                </a>
            `);

        } finally {

            if (connection) {
                connection.release();
            }
        }
    }
);

// =====================================================
// SAVE USER CLAIMS
// =====================================================

app.post(
    "/save-claims",
    async (req, res) => {

        if (
            !req.session.user ||
            normalizeRole(
                req.session.user.role
            ) !== "user"
        ) {
            return res.redirect("/");
        }

        const employeeId =
            String(
                req.session.user.employee_id || ""
            ).trim();

        if (!employeeId) {

            return res.status(400).send(`
                <h2>Save Failed</h2>

                <p>
                    Employee ID is missing.
                </p>

                <a href="/user">
                    Back to Dashboard
                </a>
            `);
        }

        let connection;

        try {

            connection =
                await db.getConnection();

            await connection.beginTransaction();

            // =================================================
            // GET ONLY THIS USER'S CLAIMS
            // =================================================

            const [claims] =
                await connection.query(
                    `
                    SELECT
                        id
                    FROM claims
                    WHERE
                        TRIM(assigned_user_id)
                        =
                        TRIM(?)
                    `,
                    [
                        employeeId
                    ]
                );

            if (claims.length === 0) {

                await connection.rollback();

                return res.status(400).send(`
                    <h2>Save Failed</h2>

                    <p>
                        No claims found for
                        Employee ID:
                        <b>${employeeId}</b>
                    </p>

                    <a href="/user">
                        Back to Dashboard
                    </a>
                `);
            }

            // =================================================
            // UPDATE EACH CLAIM
            // =================================================

            for (const claim of claims) {

                const id =
                    claim.id;

                const claimType =
                    req.body[
                        `claim_type_${id}`
                    ];

                const ilomId =
                    String(
                        req.body[
                            `ilom_id_${id}`
                        ] || ""
                    ).trim() || null;

                const approveAmountRaw =
                    req.body[
                        `approve_amount_${id}`
                    ];

                const claimStatus =
                    req.body[
                        `claim_status_${id}`
                    ] || "Pending";

                const userRemark =
                    String(
                        req.body[
                            `user_remark_${id}`
                        ] || ""
                    ).trim() || null;

                const deductionAmountRaw =
                    req.body[
                        `deduction_amount_${id}`
                    ];

                const diagnosis2 =
                    String(
                        req.body[
                            `diagnosis_2_${id}`
                        ] || ""
                    ).trim() || null;

                const interDocExe =
                    String(
                        req.body[
                            `inter_doc_exe_${id}`
                        ] || ""
                    ).trim() || null;

                // =================================================
                // CLAIM TYPE
                // =================================================

                const finalClaimType =
                    normalizeClaimType(
                        claimType
                    );

                // =================================================
                // STATUS
                // =================================================

                const finalClaimStatus =
                    normalizeStatus(
                        claimStatus
                    );

                if (
                    !VALID_STATUSES.includes(
                        finalClaimStatus
                    )
                ) {

                    throw new Error(
                        `Invalid Status '${finalClaimStatus}' for Claim ID ${id}.`
                    );
                }

                // =================================================
                // AMOUNTS
                // =================================================

                const finalApproveAmount =
                    (
                        approveAmountRaw === undefined ||
                        approveAmountRaw === null ||
                        String(
                            approveAmountRaw
                        ).trim() === ""
                    )
                        ? 0
                        : Number(
                            approveAmountRaw
                        );

                const finalDeductionAmount =
                    (
                        deductionAmountRaw === undefined ||
                        deductionAmountRaw === null ||
                        String(
                            deductionAmountRaw
                        ).trim() === ""
                    )
                        ? 0
                        : Number(
                            deductionAmountRaw
                        );

                if (
                    !Number.isFinite(
                        finalApproveAmount
                    )
                ) {

                    throw new Error(
                        `Invalid Approve Amount for Claim ID ${id}.`
                    );
                }

                if (
                    !Number.isFinite(
                        finalDeductionAmount
                    )
                ) {

                    throw new Error(
                        `Invalid Deduction Amount for Claim ID ${id}.`
                    );
                }

                // =================================================
                // UPDATE
                // =================================================

                await connection.query(
                    `
                    UPDATE claims
                    SET

                        claim_type = ?,

                        ilom_id = ?,

                        approve_amount = ?,

                        claim_status = ?,

                        user_remark = ?,

                        deduction_amount = ?,

                        diagnosis_2 = ?,

                        inter_doc_exe = ?,

                        updated_at = NOW()

                    WHERE
                        id = ?

                        AND TRIM(assigned_user_id)
                            =
                            TRIM(?)
                    `,
                    [

                        finalClaimType,

                        ilomId,

                        finalApproveAmount,

                        finalClaimStatus,

                        userRemark,

                        finalDeductionAmount,

                        diagnosis2,

                        interDocExe,

                        id,

                        employeeId
                    ]
                );
            }

            // =================================================
            // COMMIT
            // =================================================

            await connection.commit();

            console.log(
                `CLAIMS SAVED: ${employeeId}`
            );

            return res.redirect(
                "/user?saved=1"
            );

        } catch (error) {

            if (connection) {

                try {
                    await connection.rollback();
                } catch (rollbackError) {

                    console.error(
                        "ROLLBACK ERROR:",
                        rollbackError
                    );
                }
            }

            console.error(
                "SAVE CLAIMS ERROR:",
                error
            );

            return res.status(500).send(`
                <h2>
                    Save Claims Failed
                </h2>

                <pre>
${error.message}
                </pre>

                <br>

                <a href="/user">
                    Back to User Dashboard
                </a>
            `);

        } finally {

            if (connection) {
                connection.release();
            }
        }
    }
);

// =====================================================
// USER DASHBOARD
// =====================================================

app.get(
    "/user",
    async (req, res) => {

        if (
            !req.session.user ||
            normalizeRole(
                req.session.user.role
            ) !== "user"
        ) {
            return res.redirect("/");
        }

        try {

            const employeeId =
                String(
                    req.session.user.employee_id || ""
                ).trim();

            if (!employeeId) {

                return res.status(400).send(`
                    <h2>User Dashboard Error</h2>

                    <p>
                        Employee ID is missing.
                    </p>

                    <a href="/">
                        Login Again
                    </a>
                `);
            }

            // =================================================
            // USER CLAIMS
            // =================================================

            const [claims] =
                await db.query(
                    `
                    SELECT

                        c.*,

                        u.id AS employeeid,

                        u.employee_id AS employee_id,

                        u.username AS employee_name,

                        ub.uploaded_at AS uploaded_at

                    FROM claims c

                    LEFT JOIN users u
                        ON TRIM(
                            c.assigned_user_id
                        )
                        =
                        TRIM(
                            u.employee_id
                        )

                    LEFT JOIN upload_batches ub
                        ON c.upload_batch_id
                        =
                        ub.id

                    WHERE
                        TRIM(
                            c.assigned_user_id
                        )
                        =
                        TRIM(?)

                    ORDER BY
                        c.id DESC
                    `,
                    [
                        employeeId
                    ]
                );

            // =================================================
            // USER PROCESS SUMMARY
            // =================================================
            //
            // VERY IMPORTANT:
            // This is calculated AFTER SAVE
            // directly from current DB values.
            //
            // =================================================

            const [[userSummary]] =
                await db.query(
                    `
                    SELECT

                        COUNT(*) AS total,

                        SUM(
                            claim_status = 'Pending'
                        ) AS pending,

                        SUM(
                            claim_status = 'Approved'
                        ) AS approved,

                        SUM(
                            claim_status = 'Rejected'
                        ) AS rejected,

                        SUM(
                            claim_status = 'Query'
                        ) AS query_count,

                        SUM(
                            claim_status = 'Re-Query'
                        ) AS requery,

                        SUM(
                            claim_status =
                            'Query & Investigation'
                        ) AS investigation_query,

                        SUM(
                            claim_status =
                            'Investigation'
                        ) AS investigation,

                        SUM(
                            claim_status =
                            'Sent-Back'
                        ) AS sent_back,

                        SUM(
                            claim_status = 'Keep'
                        ) AS keep_count,

                        SUM(
                            claim_status =
                            'Other-Doctor/Executive'
                        ) AS other_doctor_executive,

                        SUM(
                            claim_status =
                            'ROD-Cancel'
                        ) AS rod_cancel,

                        SUM(
                            claim_status <> 'Pending'
                        ) AS total_productivity

                    FROM claims

                    WHERE
                        TRIM(
                            assigned_user_id
                        )
                        =
                        TRIM(?)
                    `,
                    [
                        employeeId
                    ]
                );

            // =================================================
            // FORMAT CLAIMS
            // =================================================

            const formattedClaims =
                claims.map(
                    claim => {

                        return {

                            ...claim,

                            formatted_claim_date:
                                formatDisplayDate(
                                    claim.claim_date
                                ),

                            formatted_claim_time:
                                formatDisplayTime(
                                    claim.claim_time
                                ),

                            formatted_upload_time:
                                formatUploadedAt(
                                    claim.uploaded_at
                                )
                        };
                    }
                );

            // =================================================
            // SUMMARY OBJECT
            // =================================================

            const processSummary = {

                total:
                    Number(
                        userSummary.total || 0
                    ),

                pending:
                    Number(
                        userSummary.pending || 0
                    ),

                approved:
                    Number(
                        userSummary.approved || 0
                    ),

                rejected:
                    Number(
                        userSummary.rejected || 0
                    ),

                query:
                    Number(
                        userSummary.query_count || 0
                    ),

                requery:
                    Number(
                        userSummary.requery || 0
                    ),

                investigationQuery:
                    Number(
                        userSummary.investigation_query || 0
                    ),

                investigation:
                    Number(
                        userSummary.investigation || 0
                    ),

                sentBack:
                    Number(
                        userSummary.sent_back || 0
                    ),

                keep:
                    Number(
                        userSummary.keep_count || 0
                    ),

                otherDoctorExecutive:
                    Number(
                        userSummary.other_doctor_executive || 0
                    ),

                rodCancel:
                    Number(
                        userSummary.rod_cancel || 0
                    ),

                totalProductivity:
                    Number(
                        userSummary.total_productivity || 0
                    )
            };

            console.log(
                "================================="
            );

            console.log(
                "USER:",
                employeeId
            );

            console.log(
                "CLAIMS:",
                formattedClaims.length
            );

            console.log(
                "USER PROCESS SUMMARY:",
                processSummary
            );

            console.log(
                "================================="
            );

            return res.render(
                "user-dashboard",
                {

                    user:
                        req.session.user,

                    claims:
                        formattedClaims,

                    saved:
                        req.query.saved === "1",

                    processSummary:
                        processSummary
                }
            );

        } catch (error) {

            console.error(
                "USER DASHBOARD ERROR:",
                error
            );

            return res.status(500).send(`
                <h2>
                    User Dashboard Error
                </h2>

                <pre>
${error.message}
                </pre>

                <br>

                <a href="/user">
                    Back to Dashboard
                </a>
            `);
        }
    }
);

// =====================================================
// USER-DASHBOARD OLD ROUTE
// =====================================================
//
// We don't need a second duplicate dashboard logic.
// Redirect it to the correct /user route.
//
// =====================================================

app.get(
    "/user-dashboard",
    (req, res) => {

        return res.redirect("/user");
    }
);

// =====================================================
// LOGOUT
// =====================================================

app.get(
    "/logout",
    (req, res) => {

        req.session.destroy(
            err => {

                if (err) {

                    console.error(
                        "LOGOUT ERROR:",
                        err
                    );
                }

                res.clearCookie(
                    "connect.sid"
                );

                return res.redirect("/");
            }
        );
    }
);

// =====================================================
// START SERVER
// =====================================================

app.listen(
    PORT,
    "0.0.0.0",
    async () => {

        console.log(
            `Server running on port ${PORT}`
        );

        await testDatabase();
    }
);