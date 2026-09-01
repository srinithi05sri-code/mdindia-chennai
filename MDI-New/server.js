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

        const ext = path
            .extname(file.originalname)
            .toLowerCase();

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

    port: Number(
        process.env.MYSQLPORT || 3306
    ),

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
        extended: true,
        limit: "20mb"
    })
);

app.use(
    express.json({
        limit: "20mb"
    })
);

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
// SAVE LOCK
// Prevent accidental simultaneous double-save
// =====================================================

const savingUsers = new Set();


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

    let value = String(status || "").trim();

    if (value === "") {
        return "Pending";
    }

    const normalized = value
        .replace(/\s+/g, " ")
        .trim();

    if (
        normalized === "SentBack" ||
        normalized === "Sent Back"
    ) {
        return "Sent-Back";
    }

    if (
        normalized === "ROD Cancel" ||
        normalized === "ROD-Cancel"
    ) {
        return "ROD-Cancel";
    }

    if (
        normalized === "Investigation&Query" ||
        normalized === "Investigation & Query"
    ) {
        return "Query & Investigation";
    }

    if (
        normalized === "OtherDoctor/Executive" ||
        normalized === "Other Doctor & Executive" ||
        normalized === "Other Doctor/Executive"
    ) {
        return "Other-Doctor/Executive";
    }

    return normalized;
}


// =====================================================
// CLAIM TYPE NORMALIZER
// =====================================================

function normalizeClaimType(value) {

    let claimType = String(value || "").trim();

    if (claimType === "") {
        return null;
    }

    const upper = claimType.toUpperCase();

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

    if (
        !VALID_CLAIM_TYPES.includes(claimType)
    ) {
        throw new Error(
            `Invalid Claim Type '${claimType}'.`
        );
    }

    return claimType;
}


// =====================================================
// EXCEL DATE CONVERTER
// =====================================================

function convertExcelDate(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {
        return null;
    }

    // Excel numeric serial

    if (
        typeof value === "number" &&
        Number.isFinite(value)
    ) {

        const date =
            XLSX.SSF.parse_date_code(value);

        if (!date) {
            return null;
        }

        const year = String(date.y);

        const month =
            String(date.m).padStart(2, "0");

        const day =
            String(date.d).padStart(2, "0");

        return `${year}-${month}-${day}`;
    }

    // JavaScript Date

    if (value instanceof Date) {

        if (isNaN(value.getTime())) {
            return null;
        }

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

    // String

    const text =
        String(value).trim();

    if (!text) {
        return null;
    }

    // YYYY-MM-DD

    if (
        /^\d{4}-\d{2}-\d{2}$/.test(text)
    ) {
        return text;
    }

    // DD-MM-YYYY / DD/MM/YYYY

    let match =
        text.match(
            /^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/
        );

    if (match) {

        const day =
            String(match[1]).padStart(2, "0");

        const month =
            String(match[2]).padStart(2, "0");

        const year =
            match[3];

        return `${year}-${month}-${day}`;
    }

    // DD/MM/YYYY with time

    match =
        text.match(
            /^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/
        );

    if (match) {

        const day =
            String(match[1]).padStart(2, "0");

        const month =
            String(match[2]).padStart(2, "0");

        const year =
            match[3];

        return `${year}-${month}-${day}`;
    }

    // Excel serial stored as string

    if (
        /^\d+(\.\d+)?$/.test(text)
    ) {

        const serial =
            Number(text);

        if (Number.isFinite(serial)) {

            const date =
                XLSX.SSF.parse_date_code(serial);

            if (date) {

                const year =
                    String(date.y);

                const month =
                    String(date.m).padStart(2, "0");

                const day =
                    String(date.d).padStart(2, "0");

                return `${year}-${month}-${day}`;
            }
        }
    }

    return null;
}


// =====================================================
// EXCEL TIME CONVERTER
// =====================================================

function convertExcelTime(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {
        return null;
    }

    // Excel decimal time

    if (
        typeof value === "number" &&
        value >= 0 &&
        value < 1
    ) {

        const totalSeconds =
            Math.round(
                value * 24 * 60 * 60
            );

        const hours =
            Math.floor(
                totalSeconds / 3600
            );

        const minutes =
            Math.floor(
                (totalSeconds % 3600) / 60
            );

        const seconds =
            totalSeconds % 60;

        return [
            String(hours).padStart(2, "0"),
            String(minutes).padStart(2, "0"),
            String(seconds).padStart(2, "0")
        ].join(":");
    }

    const text =
        String(value).trim();

    if (!text) {
        return null;
    }

    // HH:MM

    if (
        /^\d{1,2}:\d{2}$/.test(text)
    ) {
        return `${text}:00`;
    }

    // HH:MM:SS

    if (
        /^\d{1,2}:\d{2}:\d{2}$/.test(text)
    ) {
        return text;
    }

    return text;
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

app.get(
    "/",
    (req, res) => {

        res.render(
            "login",
            {
                error: null
            }
        );
    }
);


// =====================================================
// LOGIN
// =====================================================

app.post(
    "/login",
    async (req, res) => {

        const employee_id =
            String(
                req.body.employee_id || ""
            ).trim();

        const password =
            String(
                req.body.password || ""
            ).trim();

        if (
            !employee_id ||
            !password
        ) {

            return res.render(
                "login",
                {
                    error:
                        "Employee ID and Password are required"
                }
            );
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

            if (
                users.length === 0
            ) {

                return res.render(
                    "login",
                    {
                        error:
                            "Invalid Employee ID or Password"
                    }
                );
            }

            const user =
                users[0];

            const dbPassword =
                String(
                    user.password || ""
                ).trim();

            if (
                dbPassword !== password
            ) {

                return res.render(
                    "login",
                    {
                        error:
                            "Invalid Employee ID or Password"
                    }
                );
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

                return res.render(
                    "login",
                    {
                        error:
                            "Your account is inactive"
                    }
                );
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

            if (
                role === "admin"
            ) {

                return res.redirect(
                    "/admin"
                );
            }

            if (
                role === "upload"
            ) {

                return res.redirect(
                    "/upload"
                );
            }

            if (
                role === "user"
            ) {

                return res.redirect(
                    "/user"
                );
            }

            return res.render(
                "login",
                {
                    error:
                        "Invalid user role"
                }
            );

        } catch (error) {

            console.error(
                "LOGIN DATABASE ERROR:",
                error
            );

            return res.render(
                "login",
                {
                    error:
                        "Server error. Please try again."
                }
            );
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

            // ---------------------------------------------
            // USERS
            // ---------------------------------------------

            const [userList] =
                await db.query(
                    `
                    SELECT
                        id,
                        employee_id,
                        username,
                        department,
                        is_active
                    FROM users
                    WHERE LOWER(TRIM(role)) = 'user'
                    ORDER BY username
                    `
                );


            // ---------------------------------------------
            // PROCESS SUMMARY
            // ---------------------------------------------

            const [processSummary] =
                await db.query(
                    `
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
                            c.claim_status =
                            'Keep'
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
                            c.claim_status =
                            'Pending'
                        ) AS pending,

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
                    `
                );


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
// UPLOAD DASHBOARD
// IMPORTANT:
// DB COLUMN = upload_at
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

                        
        CONVERT_TZ(
            ub.uploaded_at,
            '+00:00',
            '+05:30'
        ) AS uploaded_at,

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


            return res.render(
                "upload-dashboard",
                {

                    user:
                        req.session.user,

                    uploads:
                        uploads
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

                <br>

                <a href="/upload">
                    Back to Upload
                </a>
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

            // ---------------------------------------------
            // READ WORKBOOK
            // ---------------------------------------------

            const workbook =
                XLSX.read(
                    req.file.buffer,
                    {
                        type: "buffer",
                        cellDates: false
                    }
                );

            if (
                !workbook.SheetNames ||
                workbook.SheetNames.length === 0
            ) {

                throw new Error(
                    "Excel file does not contain any sheet."
                );
            }

            const sheetName =
                workbook.SheetNames[0];

            const sheet =
                workbook.Sheets[sheetName];

            const rows =
                XLSX.utils.sheet_to_json(
                    sheet,
                    {
                        defval: "",
                        raw: true
                    }
                );

            if (
                rows.length === 0
            ) {

                throw new Error(
                    "Excel file is empty."
                );
            }


            // ---------------------------------------------
            // REQUIRED COLUMNS
            // ---------------------------------------------

            const requiredColumns = [

                "CLAIM_REF_NO",

                "INWARD_NO",

                "POLICY_NO",

                "CLAIM_AMT",

                "Vertical",

                

                "User ID",

                "User Name",

                "Claim Type",

                "Status",

                "Date",

                "Time"
            ];


          // ---------------------------------------------
// NORMALIZE EXCEL HEADERS
// ---------------------------------------------

const actualColumns = Object.keys(rows[0]);

const normalizedColumns = actualColumns.map(
    column =>
        String(column)
            .trim()
            .replace(/\s+/g, " ")
            .toLowerCase()
);

// Check required columns
const missingColumns =
    requiredColumns.filter(
        requiredColumn => {

            const normalizedRequired =
                String(requiredColumn)
                    .trim()
                    .replace(/\s+/g, " ")
                    .toLowerCase();

            return !normalizedColumns.includes(
                normalizedRequired
            );
        }
    );

if (missingColumns.length > 0) {

    throw new Error(
        "Invalid Excel Format. Missing columns: " +
        missingColumns.join(", ")
    );
}

            // ---------------------------------------------
            // CONNECTION
            // ---------------------------------------------

            connection =
                await db.getConnection();

            await connection.beginTransaction();


            // ---------------------------------------------
            // CREATE UPLOAD BATCH
            // IMPORTANT:
            // upload_at, NOT uploaded_at
            // ---------------------------------------------

            const [batchResult] =
                await connection.query(
                    `
                    INSERT INTO upload_batches
                    (
                        file_name,
                        total_claims,
                        status,
                        uploaded_at
                    )
                    VALUES
                    (
                        ?,
                        ?,
                        'ACTIVE',
                        NOW()
                    )
                    `,
                    [
                        req.file.originalname,
                        rows.length
                    ]
                );


            const batchId =
                batchResult.insertId;


            // ---------------------------------------------
            // INSERT CLAIMS
            // ---------------------------------------------

            for (
                let index = 0;
                index < rows.length;
                index++
            ) {

                const row =
                    rows[index];


                // -----------------------------------------
                // BASIC FIELDS
                // -----------------------------------------

                const claimRefNo =
                    String(
                        row["CLAIM_REF_NO"] || ""
                    ).trim();


                if (!claimRefNo) {

                    throw new Error(
                        `CLAIM_REF_NO missing at Excel row ${index + 2}.`
                    );
                }


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
                    String(
                        row["Department"] || ""
                    ).trim();


                const assignedUserId =
                    String(
                        row["User ID"] || ""
                    ).trim();


                const userName =
                    String(
                        row["User Name"] || ""
                    ).trim();


                // -----------------------------------------
                // OPTIONAL FIELDS
                // -----------------------------------------

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


                // -----------------------------------------
                // DATE
                // -----------------------------------------

                const claimDate =
                    convertExcelDate(
                        row["Date"]
                    );


                if (
                    row["Date"] !== "" &&
                    row["Date"] !== null &&
                    row["Date"] !== undefined &&
                    !claimDate
                ) {

                    throw new Error(
                        `Invalid Date for Claim ${claimRefNo}.`
                    );
                }


                // -----------------------------------------
                // TIME
                // -----------------------------------------

                const claimTime =
                    convertExcelTime(
                        row["Time"]
                    );


                // -----------------------------------------
                // CLAIM TYPE
                // -----------------------------------------

                const claimType =
                    normalizeClaimType(
                        row["Claim Type"]
                    );


                // -----------------------------------------
                // STATUS
                // -----------------------------------------

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


                // -----------------------------------------
                // INSERT
                // -----------------------------------------

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
                        ?,

                        ?, ?, ?, ?,

                        ?, ?,

                        ?, ?,

                        ?, ?, ?,

                        ?, ?, ?,

                        ?, ?, ?, ?,

                        ?, ?, ?,

                        ?, ?,

                        ?, ?, ?,

                        ?, ?,

                        ?, ?, ?,

                        ?,

                        ?, ?
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


            // ---------------------------------------------
            // COMMIT
            // ---------------------------------------------

            await connection.commit();


            console.log(
                "EXCEL UPLOAD SUCCESS:",
                req.file.originalname
            );


            return res.redirect(
                "/upload"
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
                "EXCEL UPLOAD ERROR:",
                error
            );


            return res.status(500).send(`
                <h2>Excel Upload Failed</h2>

                <pre>${error.message}</pre>

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
            String(
                req.params.id || ""
            ).trim();


        if (!batchId) {

            return res.status(400).send(
                "Invalid upload ID."
            );
        }


        let connection;


        try {

            connection =
                await db.getConnection();

            await connection.beginTransaction();


            // Delete claims belonging to batch

            await connection.query(
                `
                DELETE FROM claims
                WHERE upload_batch_id = ?
                `,
                [
                    batchId
                ]
            );


            // Mark upload deleted

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


            return res.redirect(
                "/upload"
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
                "DELETE UPLOAD ERROR:",
                error
            );


            return res.status(500).send(`
                <h2>Delete Failed</h2>

                <pre>${error.message}</pre>

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
// SAVE USER CLAIMS
// =====================================================

// ============================================================
// SAVE CLAIMS
// ============================================================

app.post("/save-claims", async (req, res) => {

    // ---------------------------------------------------------
    // LOGIN CHECK
    // ---------------------------------------------------------
    if (
        !req.session.user ||
        normalizeRole(req.session.user.role) !== "user"
    ) {
        return res.redirect("/");
    }

    const employeeId = String(
        req.session.user.employee_id || ""
    ).trim();

    if (!employeeId) {
        return res.status(400).send(`
            <h2>Save Failed</h2>
            <p>Employee ID is missing.</p>

            <a href="/user">
                Back to Dashboard
            </a>
        `);
    }

    // ---------------------------------------------------------
    // DOUBLE SAVE PROTECTION
    // ---------------------------------------------------------
    if (savingUsers.has(employeeId)) {

        return res.status(409).send(`
            <h2>Save Already In Progress</h2>

            <p>
                Your previous save request is still processing.
                Please wait a moment.
            </p>

            <a href="/user">
                Back to User Dashboard
            </a>
        `);
    }

    savingUsers.add(employeeId);

    let connection;

    try {

        // =====================================================
        // FIND CLAIM IDS FROM FORM
        // =====================================================

        const submittedKeys = Object.keys(req.body);

        const claimIds = [
            ...new Set(

                submittedKeys

                    .map(key => {

                        const match = key.match(
                            /^(?:claim_type|ilom_id|approve_amount|claim_status|user_remark|deduction_amount|diagnosis_2|inter_doc_exe)_(\d+)$/
                        );

                        return match
                            ? Number(match[1])
                            : null;
                    })

                    .filter(
                        id =>
                            Number.isInteger(id) &&
                            id > 0
                    )
            )
        ];

        if (claimIds.length === 0) {

            return res.status(400).send(`
                <h2>Save Failed</h2>

                <p>
                    No claim changes were submitted.
                </p>

                <a href="/user">
                    Back to Dashboard
                </a>
            `);
        }

        console.log(
            "SAVE CLAIM IDS:",
            claimIds
        );

        // =====================================================
        // DB CONNECTION
        // =====================================================

        connection = await db.getConnection();

        await connection.beginTransaction();

        // =====================================================
        // UPDATE CLAIMS
        // =====================================================

        for (const id of claimIds) {

            // -------------------------------------------------
            // SECURITY CHECK
            // Claim must belong to logged-in employee
            // -------------------------------------------------

            const [claimRows] =
                await connection.query(
                    `
                    SELECT id
                    FROM claims
                    WHERE id = ?
                    AND TRIM(assigned_user_id) = TRIM(?)
                    LIMIT 1
                    `,
                    [
                        id,
                        employeeId
                    ]
                );

            if (claimRows.length === 0) {

                throw new Error(
                    `Claim ID ${id} does not belong to this user.`
                );
            }

            // =================================================
            // FORM VALUES
            // =================================================

            const claimTypeRaw =
                req.body[`claim_type_${id}`];

            const ilomId =
                String(
                    req.body[`ilom_id_${id}`] || ""
                ).trim() || null;

            const approveAmountRaw =
                req.body[`approve_amount_${id}`];

            const claimStatusRaw =
                req.body[`claim_status_${id}`];

            const userRemark =
                String(
                    req.body[`user_remark_${id}`] || ""
                ).trim() || null;

            const deductionAmountRaw =
                req.body[`deduction_amount_${id}`];

            const diagnosis2 =
                String(
                    req.body[`diagnosis_2_${id}`] || ""
                ).trim() || null;

            const interDocExe =
                String(
                    req.body[`inter_doc_exe_${id}`] || ""
                ).trim() || null;

            // =================================================
            // CLAIM TYPE
            // =================================================

            const finalClaimType =
                normalizeClaimType(
                    claimTypeRaw
                );

            // =================================================
            // STATUS
            // =================================================

            const finalClaimStatus =
                normalizeStatus(
                    claimStatusRaw
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
            // APPROVE AMOUNT
            // =================================================

            let finalApproveAmount = 0;

            if (
                approveAmountRaw !== undefined &&
                approveAmountRaw !== null &&
                String(approveAmountRaw).trim() !== ""
            ) {

                finalApproveAmount =
                    Number(
                        String(approveAmountRaw)
                            .replace(/,/g, "")
                            .trim()
                    );
            }

            if (
                !Number.isFinite(
                    finalApproveAmount
                )
            ) {

                throw new Error(
                    `Invalid Approve Amount for Claim ID ${id}.`
                );
            }

            // =================================================
            // DEDUCTION AMOUNT
            // =================================================

            let finalDeductionAmount = 0;

            if (
                deductionAmountRaw !== undefined &&
                deductionAmountRaw !== null &&
                String(deductionAmountRaw).trim() !== ""
            ) {

                finalDeductionAmount =
                    Number(
                        String(deductionAmountRaw)
                            .replace(/,/g, "")
                            .trim()
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
                    updated_at =
                        CONVERT_TZ(
                            UTC_TIMESTAMP(),
                            '+00:00',
                            '+05:30'
                        )

                WHERE
                    id = ?

                AND
                    TRIM(assigned_user_id) = TRIM(?)
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

        // =====================================================
        // COMMIT
        // =====================================================

        await connection.commit();

        console.log(
            `CLAIMS SAVED SUCCESSFULLY: ${employeeId}`
        );

        // =====================================================
        // REDIRECT
        // =====================================================

        return res.redirect(
            "/user?saved=1"
        );

    } catch (error) {

        // =====================================================
        // ROLLBACK
        // =====================================================

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
            <h2>Save Claims Failed</h2>

            <p>
                ${error.message}
            </p>

            <br>

            <a href="/user">
                Back to User Dashboard
            </a>
        `);

    } finally {

        if (connection) {
            connection.release();
        }

        savingUsers.delete(employeeId);
    }
});


// ============================================================
// USER DASHBOARD
// ============================================================

app.get("/user", async (req, res) => {

    // ---------------------------------------------------------
    // LOGIN CHECK
    // ---------------------------------------------------------

    if (
        !req.session.user ||
        normalizeRole(req.session.user.role) !== "user"
    ) {
        return res.redirect("/");
    }

    try {

        const employeeId = String(
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

        // =====================================================
        // CLAIMS
        // =====================================================

        const [claims] = await db.query(
            `
            SELECT

                c.*,

                u.id AS employeeid,

                u.employee_id AS employee_id,

                u.username AS employee_name,

                ub.uploaded_at AS uploaded_at

            FROM claims c

            LEFT JOIN users u

                ON TRIM(c.assigned_user_id)
                =
                TRIM(u.employee_id)

            LEFT JOIN upload_batches ub

                ON c.upload_batch_id = ub.id

            WHERE
                TRIM(c.assigned_user_id)
                =
                TRIM(?)

            ORDER BY
                c.id DESC
            `,
            [
                employeeId
            ]
        );

        // =====================================================
        // USER SUMMARY
        // =====================================================

        const [[userSummary]] = await db.query(
            `
            SELECT

                COUNT(*) AS total_allocated,

                COALESCE(
                    SUM(
                        claim_status = 'Pending'
                    ),
                    0
                ) AS pending,

                COALESCE(
                    SUM(
                        claim_status = 'Approved'
                    ),
                    0
                ) AS approved,

                COALESCE(
                    SUM(
                        claim_status = 'Rejected'
                    ),
                    0
                ) AS rejected,

                COALESCE(
                    SUM(
                        claim_status = 'Query'
                    ),
                    0
                ) AS query_count,

                COALESCE(
                    SUM(
                        claim_status = 'Re-Query'
                    ),
                    0
                ) AS requery,

                COALESCE(
                    SUM(
                        claim_status =
                        'Query & Investigation'
                    ),
                    0
                ) AS investigation_query,

                COALESCE(
                    SUM(
                        claim_status =
                        'Investigation'
                    ),
                    0
                ) AS investigation,

                COALESCE(
                    SUM(
                        claim_status =
                        'Sent-Back'
                    ),
                    0
                ) AS sent_back,

                COALESCE(
                    SUM(
                        claim_status = 'Keep'
                    ),
                    0
                ) AS keep_count,

                COALESCE(
                    SUM(
                        claim_status =
                        'Other-Doctor/Executive'
                    ),
                    0
                ) AS other_doctor_executive,

                COALESCE(
                    SUM(
                        claim_status =
                        'ROD-Cancel'
                    ),
                    0
                ) AS rod_cancel,

                COALESCE(
                    SUM(
                        claim_status IN (
                            'Approved',
                            'Rejected',
                            'Query',
                            'Re-Query',
                            'Query & Investigation',
                            'Investigation',
                            'Sent-Back',
                            'Keep',
                            'Other-Doctor/Executive',
                            'ROD-Cancel'
                        )
                    ),
                    0
                ) AS total_productivity

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

        // =====================================================
        // PLATFORM SUMMARY
        // =====================================================

        const [platformSummary] =
            await db.query(
                `
                SELECT

                    COALESCE(
                        platform,
                        '-'
                    ) AS platform,

                    COUNT(*) AS total_allocated,

                    COALESCE(
                        SUM(
                            claim_status = 'Approved'
                        ),
                        0
                    ) AS approved,

                    COALESCE(
                        SUM(
                            claim_status = 'Rejected'
                        ),
                        0
                    ) AS rejected,

                    COALESCE(
                        SUM(
                            claim_status = 'Query'
                        ),
                        0
                    ) AS query_count,

                    COALESCE(
                        SUM(
                            claim_status = 'Re-Query'
                        ),
                        0
                    ) AS requery,

                    COALESCE(
                        SUM(
                            claim_status =
                            'Query & Investigation'
                        ),
                        0
                    ) AS investigation_query,

                    COALESCE(
                        SUM(
                            claim_status =
                            'Investigation'
                        ),
                        0
                    ) AS investigation,

                    COALESCE(
                        SUM(
                            claim_status =
                            'Sent-Back'
                        ),
                        0
                    ) AS sent_back,

                    COALESCE(
                        SUM(
                            claim_status = 'Keep'
                        ),
                        0
                    ) AS keep_count,

                    COALESCE(
                        SUM(
                            claim_status =
                            'Other-Doctor/Executive'
                        ),
                        0
                    ) AS other_doctor_executive,

                    COALESCE(
                        SUM(
                            claim_status = 'Pending'
                        ),
                        0
                    ) AS pending,

                    COALESCE(
                        SUM(
                            claim_status =
                            'ROD-Cancel'
                        ),
                        0
                    ) AS rod_cancel,

                    COALESCE(
                        SUM(
                            claim_status IN (
                                'Approved',
                                'Rejected',
                                'Query',
                                'Re-Query',
                                'Query & Investigation',
                                'Investigation',
                                'Sent-Back',
                                'Keep',
                                'Other-Doctor/Executive',
                                'ROD-Cancel'
                            )
                        ),
                        0
                    ) AS total_productivity

                FROM claims

                WHERE
                    TRIM(assigned_user_id)
                    =
                    TRIM(?)

                GROUP BY
                    platform

                ORDER BY
                    platform
                `,
                [
                    employeeId
                ]
            );

        // =====================================================
        // FORMAT CLAIMS
        // =====================================================

        const formattedClaims =
            claims.map(claim => {

                let formattedDate = "-";
                let formattedTime = "-";
                let formattedUploadedAt = "-";

                // -------------------------------------------------
                // CLAIM DATE
                // -------------------------------------------------

                if (claim.claim_date) {

                    const date =
                        new Date(
                            claim.claim_date
                        );

                    if (
                        !isNaN(
                            date.getTime()
                        )
                    ) {

                        formattedDate =
                            date.toLocaleDateString(
                                "en-IN",
                                {
                                    day: "2-digit",
                                    month: "2-digit",
                                    year: "numeric"
                                }
                            );
                    }
                }

                // -------------------------------------------------
                // CLAIM TIME
                // -------------------------------------------------

                if (claim.claim_time) {

                    const timeValue =
                        String(
                            claim.claim_time
                        ).trim();

                    if (
                        /^\d{2}:\d{2}:\d{2}$/
                            .test(timeValue)
                    ) {

                        const [
                            hour,
                            minute
                        ] =
                            timeValue.split(":");

                        let h =
                            Number(hour);

                        const suffix =
                            h >= 12
                                ? "PM"
                                : "AM";

                        h =
                            h % 12 || 12;

                        formattedTime =
                            `${String(h).padStart(2, "0")}:${minute} ${suffix}`;

                    } else {

                        formattedTime =
                            timeValue;
                    }
                }

                // -------------------------------------------------
                // UPLOAD TIME
                // -------------------------------------------------

                if (claim.uploaded_at) {

                    const uploadedDate =
                        new Date(
                            claim.uploaded_at
                        );

                    if (
                        !isNaN(
                            uploadedDate.getTime()
                        )
                    ) {

                        formattedUploadedAt =
                            uploadedDate.toLocaleString(
                                "en-IN",
                                {
                                    day: "2-digit",
                                    month: "2-digit",
                                    year: "numeric",
                                    hour: "2-digit",
                                    minute: "2-digit",
                                    second: "2-digit",
                                    hour12: true
                                }
                            );
                    }
                }

                return {

                    ...claim,

                    formatted_claim_date:
                        formattedDate,

                    formatted_claim_time:
                        formattedTime,

                    formatted_uploaded_at:
                        formattedUploadedAt
                };
            });

        // =====================================================
        // PROCESS SUMMARY
        // =====================================================

        const processSummary = {

            totalAllocated:
                Number(
                    userSummary.total_allocated || 0
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
                ),

            // IMPORTANT
            // EJS uses processSummary.platformSummary

            platformSummary:
                platformSummary
        };

        // =====================================================
        // SAVED TIME
        // =====================================================

        let savedAt = null;

        if (
            req.query.saved === "1"
        ) {

            savedAt =
                new Date().toLocaleString(
                    "en-IN",
                    {
                        day: "2-digit",
                        month: "2-digit",
                        year: "numeric",
                        hour: "2-digit",
                        minute: "2-digit",
                        second: "2-digit",
                        hour12: true
                    }
                );
        }

        // =====================================================
        // RENDER
        // =====================================================

        return res.render(
            "user-dashboard",
            {

                user:
                    req.session.user,

                claims:
                    formattedClaims,

                saved:
                    req.query.saved === "1",

                savedAt:
                    savedAt,

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
            <h2>User Dashboard Error</h2>

            <pre>${error.message}</pre>

            <br>

            <a href="/user">
                Back to Dashboard
            </a>
        `);
    }
});


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
                    Employee ID, Username and Password are required.
                </p>

                <a href="/admin">
                    Back to Admin
                </a>
            `);
        }


        try {

            // ---------------------------------------------
            // CHECK EMPLOYEE ID
            // ---------------------------------------------

            const [employeeExists] =
                await db.query(
                    `
                    SELECT
                        id
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


            if (
                employeeExists.length > 0
            ) {

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


            // ---------------------------------------------
            // CHECK USERNAME
            // ---------------------------------------------

            const [usernameExists] =
                await db.query(
                    `
                    SELECT
                        id
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


            if (
                usernameExists.length > 0
            ) {

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


            // ---------------------------------------------
            // INSERT USER
            // ---------------------------------------------

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


            return res.redirect(
                "/admin"
            );

        } catch (error) {

            console.error(
                "CREATE USER ERROR:",
                error
            );


            return res.status(500).send(`
                <h2>Create User Failed</h2>

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


        if (
            !oldUserId ||
            !newUserId
        ) {

            return res.status(400).send(`
                <h2>Invalid User Selection</h2>

                <a href="/admin">
                    Back to Admin
                </a>
            `);
        }


        if (
            oldUserId === newUserId
        ) {

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


            // ---------------------------------------------
            // OLD USER
            // ---------------------------------------------

            const [oldUser] =
                await connection.query(
                    `
                    SELECT
                        employee_id
                    FROM users
                    WHERE id = ?
                    AND LOWER(TRIM(role)) = 'user'
                    LIMIT 1
                    `,
                    [
                        oldUserId
                    ]
                );


            if (
                oldUser.length === 0
            ) {

                throw new Error(
                    "Old user not found."
                );
            }


            const oldEmployeeId =
                String(
                    oldUser[0].employee_id
                ).trim();


            // ---------------------------------------------
            // NEW USER
            // ---------------------------------------------

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


            if (
                newUser.length === 0
            ) {

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


            // ---------------------------------------------
            // REASSIGN CLAIMS
            // ---------------------------------------------

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


            return res.redirect(
                "/admin"
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
                "REASSIGN ERROR:",
                error
            );


            return res.status(500).send(`
                <h2>Reassignment Failed</h2>

                <pre>${error.message}</pre>

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
app.get(
    "/admin/download-productivity",
    async (req, res) => {

        if (
            !req.session.user ||
            normalizeRole(req.session.user.role) !== "admin"
        ) {
            return res.redirect("/");
        }

        try {

            const [rows] = await db.query(`
                SELECT

                    COALESCE(c.platform, '-') AS platform,

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

                    SUM(c.claim_status = 'Approved') AS approved,

                    SUM(c.claim_status = 'Rejected') AS rejected,

                    SUM(c.claim_status = 'Query') AS query_count,

                    SUM(c.claim_status = 'Re-Query') AS requery,

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
                        c.claim_status =
                        'Keep'
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
                        c.claim_status =
                        'Pending'
                    ) AS pending,

                    SUM(
                        c.claim_status <> 'Pending'
                    ) AS total_productivity

                FROM claims c

                LEFT JOIN users u
                    ON TRIM(c.assigned_user_id)
                    =
                    TRIM(u.employee_id)

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

            console.log(
                "PROCESS SUMMARY ROW COUNT:",
                rows.length
            );

            const excelData = rows.map(row => ({

                "Platform":
                    row.platform || "",

                "Employee ID":
                    row.employee_id || "",

                "User Name":
                    row.user_name || "",

                "Total Allocated":
                    Number(row.total_allocated || 0),

                "Approved":
                    Number(row.approved || 0),

                "Rejected":
                    Number(row.rejected || 0),

                "Query":
                    Number(row.query_count || 0),

                "Re-Query":
                    Number(row.requery || 0),

                "Query + Investigation":
                    Number(row.investigation_query || 0),

                "Total Productivity":
                    Number(row.total_productivity || 0),

                "Investigation":
                    Number(row.investigation || 0),

                "Sent Back":
                    Number(row.sent_back || 0),

                "Keep":
                    Number(row.keep_count || 0),

                "Other Doctor & Executive":
                    Number(row.other_doctor_executive || 0),

                "ROD-Cancel":
                    Number(row.rod_cancel || 0),

                "Pending":
                    Number(row.pending || 0)
            }));

            const workbook =
                XLSX.utils.book_new();

            const worksheet =
                XLSX.utils.json_to_sheet(
                    excelData
                );

            worksheet["!cols"] = [
                { wch: 15 },
                { wch: 18 },
                { wch: 20 },
                { wch: 18 },
                { wch: 12 },
                { wch: 12 },
                { wch: 12 },
                { wch: 12 },
                { wch: 25 },
                { wch: 20 },
                { wch: 18 },
                { wch: 15 },
                { wch: 12 },
                { wch: 28 },
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
                <h2>Process Summary Download Failed</h2>
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
                await db.query(
                    `
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
                            c.claim_status =
                            'Keep'
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
                            c.claim_status =
                            'Pending'
                        ) AS pending,

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
                    `
                );


            const workbook =
                XLSX.utils.book_new();


            const excelData =
                rows.map(
                    row => ({

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
                    })
                );


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


            return res.send(
                buffer
            );

        } catch (error) {

            console.error(
                "PROCESS SUMMARY DOWNLOAD ERROR:",
                error
            );


            return res.status(500).send(`
                <h2>Process Summary Download Error</h2>

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

app.get(
    "/admin/download-claims",
    async (req, res) => {

        if (
            !req.session.user ||
            normalizeRole(req.session.user.role) !== "admin"
        ) {
            return res.redirect("/");
        }

        const fromDate = String(
            req.query.fromDate || ""
        ).trim();

        const toDate = String(
            req.query.toDate || ""
        ).trim();

        if (!fromDate || !toDate) {
            return res.status(400).send(
                "From Date and To Date are required."
            );
        }

        try {

            console.log("========== DUMP DOWNLOAD ==========");
            console.log("FROM DATE:", fromDate);
            console.log("TO DATE:", toDate);

           const [rows] = await db.query(
    `
    SELECT 
        c.*,

        ub.file_name,

       CONVERT_TZ(
            ub.uploaded_at,
            '+00:00',
            '+05:30'
        ) AS upload_date_time

    FROM claims c

    LEFT JOIN upload_batches ub
        ON c.upload_batch_id = ub.id

    WHERE
        DATE(ub.uploaded_at)
        BETWEEN ?
        AND ?

    ORDER BY
        c.id DESC
    `,
    [
        fromDate,
        toDate
    ]
);

            console.log(
                "DUMP ROW COUNT:",
                rows.length
            );

            const workbook =
                XLSX.utils.book_new();

            const excelData = rows.map(row => ({

                "CLAIM_REF_NO":
                    row.claim_ref_no || "",

                "INWARD_NO":
                    row.inward_no || "",

                "POLICY_NO":
                    row.policy_no || "",

                "CLAIM_AMT":
                    row.claim_amount || 0,

                "Vertical":
                    row.vertical || "",

                "Department":
                    row.department || "",

                "User ID":
                    row.assigned_user_id || "",

                "User Name":
                    row.user_name || "",

                "lot":
                    row.lot || "",

                "platform":
                    row.platform || "",

                "Claim Type":
                    row.claim_type || "",

                "Status":
                    row.claim_status || "",

                "Date":
                    row.claim_date || "",

                "Time":
                    row.claim_time || "",

                "Today Status":
                    row.today_status || "",

                "I3 Status":
                    row.i3_status || "",

                "Full qc":
                    row.full_qc || "",

                "RELATION":
                    row.relation || "",

                "HNF":
                    row.hnf || "",

                "ILOM ID":
                    row.ilom_id || "",

                "Approve AMT":
                    row.approve_amount || 0,

                "Remark":
                    row.user_remark || "",

                "Deduction AMT":
                    row.deduction_amount || 0,

                "Diagnosis 2":
                    row.diagnosis_2 || "",

                "inter. Doc & Exe":
                    row.inter_doc_exe || "",

               "Upload Date & Time":
         row.upload_date_time
        ? new Date(row.upload_date_time).toLocaleString(
            "en-IN",
            {
                timeZone: "Asia/Kolkata",
                day: "2-digit",
                month: "2-digit",
                year: "numeric",
                hour: "2-digit",
                minute: "2-digit",
                second: "2-digit",
                hour12: false
            }
        )
        : "",

    "Upload File":
        row.file_name || ""

            }));

            console.log(
                "DUMP EXCEL GENERATED:",
                excelData.length,
                "rows"
            );

            const worksheet =
                XLSX.utils.json_to_sheet(
                    excelData
                );

            worksheet["!cols"] = [

                { wch: 20 }, // CLAIM_REF_NO
                { wch: 18 }, // INWARD_NO
                { wch: 18 }, // POLICY_NO
                { wch: 15 }, // CLAIM_AMT
                { wch: 15 }, // Vertical
                { wch: 18 }, // Department
                { wch: 15 }, // User ID
                { wch: 20 }, // User Name
                { wch: 15 }, // lot
                { wch: 15 }, // platform
                { wch: 15 }, // Claim Type
                { wch: 25 }, // Status
                { wch: 15 }, // Date
                { wch: 15 }, // Time
                { wch: 20 }, // Today Status
                { wch: 20 }, // I3 Status
                { wch: 15 }, // Full qc
                { wch: 15 }, // RELATION
                { wch: 15 }, // HNF
                { wch: 18 }, // ILOM ID
                { wch: 15 }, // Approve AMT
                { wch: 30 }, // Remark
                { wch: 18 }, // Deduction AMT
                { wch: 25 }, // Diagnosis 2
                { wch: 25 }, // inter. Doc & Exe
                { wch: 22 }, // Upload Date & Time
                { wch: 30 }  // Upload File

            ];

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
                "attachment; filename=updated-claims.xlsx"
            );

            res.setHeader(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            );

            return res.send(buffer);

        } catch (error) {

            console.error(
                "DOWNLOAD CLAIMS ERROR:",
                error
            );

            return res.status(500).send(`
                <h2>Download Claims Failed</h2>

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
// GLOBAL MULTER / UPLOAD ERROR
// =====================================================

app.use(
    (err, req, res, next) => {

        if (
            err instanceof multer.MulterError
        ) {

            return res.status(400).send(`
                <h2>Excel Upload Error</h2>

                <pre>${err.message}</pre>

                <br>

                <a href="/upload">
                    Back to Upload
                </a>
            `);
        }


        if (
            err &&
            err.message &&
            err.message.includes(
                "Only Excel files"
            )
        ) {

            return res.status(400).send(`
                <h2>Invalid File</h2>

                <pre>${err.message}</pre>

                <br>

                <a href="/upload">
                    Back to Upload
                </a>
            `);
        }


        console.error(
            "GLOBAL ERROR:",
            err
        );


        return res.status(500).send(`
            <h2>Server Error</h2>

            <pre>${err.message || "Unknown error"}</pre>

            <br>

            <a href="/">
                Back to Login
            </a>
        `);
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