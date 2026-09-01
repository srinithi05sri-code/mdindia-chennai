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
// MULTER
// =====================================================

const upload = multer({
    storage: multer.memoryStorage(),

    limits: {
        fileSize: 10 * 1024 * 1024
    },

    fileFilter: (req, file, cb) => {

        const ext =
            path.extname(
                file.originalname
            ).toLowerCase();

        if (
            ext === ".xlsx" ||
            ext === ".xls"
        ) {
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
// MYSQL
// =====================================================

const db = mysql.createPool({
    host: process.env.DB_HOST || "localhost",
    user: process.env.DB_USER || "root",
    password: process.env.DB_PASSWORD || "1234",
    database:
        process.env.DB_NAME ||
        "claim_management",

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
    path.join(
        __dirname,
        "views"
    )
);

app.use(
    express.urlencoded({
        extended: true
    })
);

app.use(
    express.json()
);

app.use(
    express.static(
        path.join(
            __dirname,
            "public"
        )
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
            maxAge:
                1000 * 60 * 60
        }
    })
);

// =====================================================
// CONSTANTS
// =====================================================

const VALID_CLAIM_TYPES = [
    "IPD",
    "OPD",
    "PrePost"
];

const VALID_STATUSES = [
    "Pending",
    "Approved",
    "Rejected",
    "Query",
    "Re-Query",
    "Investigation&Query",
    "Investigation",
    "SentBack",
    "Keep",
    "OtherDoctor/Executive",
    "ROD Cancel"
];

// =====================================================
// DOUBLE SAVE PROTECTION
// =====================================================

const savingUsers =
    new Set();

// =====================================================
// HELPERS
// =====================================================

function normalizeRole(role) {

    return String(
        role || ""
    )
        .trim()
        .toLowerCase();
}


// -----------------------------------------------------
// CLAIM TYPE
// -----------------------------------------------------

function normalizeClaimType(value) {

    const valueText =
        String(
            value || ""
        )
            .trim();

    if (!valueText) {
        return null;
    }

    const upper =
        valueText.toUpperCase();

    if (
        upper === "IPD" ||
        upper === "INPATIENT"
    ) {
        return "IPD";
    }

    if (
        upper === "OPD" ||
        upper === "OUTPATIENT"
    ) {
        return "OPD";
    }

    if (
        upper === "PREPOST" ||
        upper === "PRE POST"
    ) {
        return "PrePost";
    }

    return valueText;
}


// -----------------------------------------------------
// STATUS
// -----------------------------------------------------

function normalizeStatus(value) {

    let status =
        String(
            value || ""
        )
            .trim();

    if (!status) {
        return "Pending";
    }

    const lower =
        status.toLowerCase();

    const map = {

        "pending":
            "Pending",

        "approved":
            "Approved",

        "rejected":
            "Rejected",

        "query":
            "Query",

        "re-query":
            "Re-Query",

        "requery":
            "Re-Query",

        "investigation&query":
            "Investigation&Query",

        "investigation & query":
            "Investigation&Query",

        "query & investigation":
            "Investigation&Query",

        "investigation":
            "Investigation",

        "sentback":
            "SentBack",

        "sent-back":
            "SentBack",

        "keep":
            "Keep",

        "otherdoctor/executive":
            "OtherDoctor/Executive",

        "other-doctor/executive":
            "OtherDoctor/Executive",

        "rod cancel":
            "ROD Cancel",

        "rod-cancel":
            "ROD Cancel"
    };

    return (
        map[lower] ||
        status
    );
}


// -----------------------------------------------------
// EXCEL HEADER NORMALIZATION
// -----------------------------------------------------

function normalizeHeader(value) {

    return String(
        value || ""
    )
        .replace(/\u00A0/g, " ")
        .trim()
        .replace(/\s+/g, " ")
        .toLowerCase();
}


// -----------------------------------------------------
// EXCEL ROW NORMALIZATION
// -----------------------------------------------------

function normalizeExcelRows(rows) {

    return rows.map(row => {

        const normalized = {};

        Object.keys(row).forEach(
            originalKey => {

                const key =
                    normalizeHeader(
                        originalKey
                    );

                normalized[key] =
                    row[originalKey];
            }
        );

        return normalized;
    });
}


// -----------------------------------------------------
// GET EXCEL VALUE
// -----------------------------------------------------

function excelValue(
    row,
    ...names
) {

    for (
        const name of names
    ) {

        const key =
            normalizeHeader(
                name
            );

        if (
            Object.prototype.hasOwnProperty.call(
                row,
                key
            )
        ) {

            return row[key];
        }
    }

    return "";
}


// -----------------------------------------------------
// NUMBER
// -----------------------------------------------------

function parseExcelNumber(value) {

    if (
        value === undefined ||
        value === null ||
        String(value).trim() === ""
    ) {
        return 0;
    }

    const number =
        Number(
            String(value)
                .replace(/,/g, "")
                .trim()
        );

    if (
        !Number.isFinite(number)
    ) {
        return 0;
    }

    return number;
}


// -----------------------------------------------------
// DATE
// -----------------------------------------------------

function convertExcelDate(value) {

    if (
        value === undefined ||
        value === null ||
        String(value).trim() === ""
    ) {
        return null;
    }

    // Excel serial date
    if (
        typeof value === "number"
    ) {

        const date =
            XLSX.SSF.parse_date_code(
                value
            );

        if (!date) {
            return null;
        }

        return new Date(
            date.y,
            date.m - 1,
            date.d
        );
    }

    const text =
        String(value).trim();

    // DD/MM/YYYY
    let match =
        text.match(
            /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/
        );

    if (match) {

        const day =
            Number(match[1]);

        const month =
            Number(match[2]) - 1;

        const year =
            Number(match[3]);

        return new Date(
            year,
            month,
            day
        );
    }

    const date =
        new Date(text);

    if (
        !isNaN(
            date.getTime()
        )
    ) {
        return date;
    }

    return null;
}


// -----------------------------------------------------
// TIME
// -----------------------------------------------------

function convertExcelTime(value) {

    if (
        value === undefined ||
        value === null ||
        String(value).trim() === ""
    ) {
        return null;
    }

    if (
        typeof value === "number"
    ) {

        const totalSeconds =
            Math.round(
                value * 24 * 60 * 60
            );

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

        return [
            String(hour).padStart(2, "0"),
            String(minute).padStart(2, "0"),
            String(second).padStart(2, "0")
        ].join(":");
    }

    const text =
        String(value).trim();

    const match =
        text.match(
            /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i
        );

    if (!match) {
        return text;
    }

    let hour =
        Number(match[1]);

    const minute =
        Number(match[2]);

    const second =
        Number(match[3] || 0);

    const ampm =
        match[4];

    if (
        ampm
    ) {

        if (
            ampm.toUpperCase() === "PM" &&
            hour !== 12
        ) {
            hour += 12;
        }

        if (
            ampm.toUpperCase() === "AM" &&
            hour === 12
        ) {
            hour = 0;
        }
    }

    return [
        String(hour).padStart(2, "0"),
        String(minute).padStart(2, "0"),
        String(second).padStart(2, "0")
    ].join(":");
}


// =====================================================
// DATABASE TEST
// =====================================================

async function testDatabase() {

    try {

        const connection =
            await db.getConnection();

        console.log(
            "MySQL Connected Successfully"
        );

        connection.release();

    } catch (error) {

        console.error(
            "MySQL Connection Failed:",
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

        const username =
            String(
                req.body.username || ""
            ).trim();

        const password =
            String(
                req.body.password || ""
            );

        try {

            const [users] =
                await db.query(
                    `
                    SELECT *
                    FROM users
                    WHERE
                        (
                            username = ?
                            OR employee_id = ?
                        )
                        AND password = ?
                        AND is_active = TRUE
                    LIMIT 1
                    `,
                    [
                        username,
                        username,
                        password
                    ]
                );

            if (
                users.length === 0
            ) {

                return res.render(
                    "login",
                    {
                        error:
                            "Invalid username or password"
                    }
                );
            }

            const user =
                users[0];

            req.session.user = {

                id:
                    user.id,

                employee_id:
                    user.employee_id,

                username:
                    user.username,

                role:
                    normalizeRole(
                        user.role
                    )
            };

            if (
                normalizeRole(
                    user.role
                ) === "admin"
            ) {
                return res.redirect(
                    "/admin"
                );
            }

            if (
                normalizeRole(
                    user.role
                ) === "upload"
            ) {
                return res.redirect(
                    "/upload"
                );
            }

            if (
                normalizeRole(
                    user.role
                ) === "user"
            ) {
                return res.redirect(
                    "/user"
                );
            }

            return res.redirect("/");

        } catch (error) {

            console.error(
                "LOGIN ERROR:",
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

            // =================================================
            // LIVE PROCESS SUMMARY
            // =================================================

            const [[summary]] =
                await db.query(
                    `
                    SELECT

                        COUNT(*) AS total,

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
                                'Investigation&Query'
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
                                'SentBack'
                            ),
                            0
                        ) AS sent_back,

                        COALESCE(
                            SUM(
                                claim_status =
                                'Keep'
                            ),
                            0
                        ) AS keep_count,

                        COALESCE(
                            SUM(
                                claim_status =
                                'OtherDoctor/Executive'
                            ),
                            0
                        ) AS other_doctor_executive,

                        COALESCE(
                            SUM(
                                claim_status =
                                'ROD Cancel'
                            ),
                            0
                        ) AS rod_cancel,

                        COALESCE(
                            SUM(
                                claim_status <> 'Pending'
                            ),
                            0
                        ) AS total_productivity

                    FROM claims
                    `
                );


            // =================================================
            // USER LIST
            // =================================================

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
                    WHERE
                        LOWER(
                            TRIM(role)
                        ) = 'user'
                    ORDER BY username
                    `
                );


            // =================================================
            // PLATFORM / USER SUMMARY
            // =================================================

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

                        COALESCE(
                            SUM(
                                c.claim_status =
                                'Approved'
                            ),
                            0
                        ) AS approved,

                        COALESCE(
                            SUM(
                                c.claim_status =
                                'Rejected'
                            ),
                            0
                        ) AS rejected,

                        COALESCE(
                            SUM(
                                c.claim_status =
                                'Query'
                            ),
                            0
                        ) AS query_count,

                        COALESCE(
                            SUM(
                                c.claim_status =
                                'Re-Query'
                            ),
                            0
                        ) AS requery,

                        COALESCE(
                            SUM(
                                c.claim_status =
                                'Investigation&Query'
                            ),
                            0
                        ) AS investigation_query,

                        COALESCE(
                            SUM(
                                c.claim_status =
                                'Investigation'
                            ),
                            0
                        ) AS investigation,

                        COALESCE(
                            SUM(
                                c.claim_status =
                                'SentBack'
                            ),
                            0
                        ) AS sent_back,

                        COALESCE(
                            SUM(
                                c.claim_status =
                                'Keep'
                            ),
                            0
                        ) AS keep_count,

                        COALESCE(
                            SUM(
                                c.claim_status =
                                'OtherDoctor/Executive'
                            ),
                            0
                        ) AS other_doctor_executive,

                        COALESCE(
                            SUM(
                                c.claim_status =
                                'ROD Cancel'
                            ),
                            0
                        ) AS rod_cancel,

                        COALESCE(
                            SUM(
                                c.claim_status =
                                'Pending'
                            ),
                            0
                        ) AS pending,

                        COALESCE(
                            SUM(
                                c.claim_status <>
                                'Pending'
                            ),
                            0
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


            const counts = {

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

                totalProductivity:
                    Number(
                        summary.total_productivity || 0
                    )
            };


            return res.render(
                "admin-dashboard",
                {

                    user:
                        req.session.user,

                    counts:
                        counts,

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

        const username =
            String(
                req.body.username || ""
            ).trim();

        const password =
            String(
                req.body.password || ""
            );

        const employeeId =
            String(
                req.body.employee_id || ""
            ).trim();

        const department =
            String(
                req.body.department || ""
            ).trim();

        try {

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
                    department
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
                <h2>Failed to Create User</h2>
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

            return res.status(400).send(
                "Please select both users."
            );
        }

        if (
            oldUserId === newUserId
        ) {

            return res.status(400).send(
                "Leaving user and new user cannot be the same."
            );
        }

        let connection;

        try {

            connection =
                await db.getConnection();

            await connection.beginTransaction();


            const [newUsers] =
                await connection.query(
                    `
                    SELECT
                        employee_id,
                        username
                    FROM users
                    WHERE
                        id = ?
                        AND LOWER(TRIM(role)) = 'user'
                        AND is_active = TRUE
                    LIMIT 1
                    `,
                    [
                        newUserId
                    ]
                );

            if (
                newUsers.length === 0
            ) {

                throw new Error(
                    "New user not found or inactive."
                );
            }

            const newEmployeeId =
                String(
                    newUsers[0].employee_id
                ).trim();

            const newUsername =
                newUsers[0].username;


            const [oldUsers] =
                await connection.query(
                    `
                    SELECT
                        employee_id
                    FROM users
                    WHERE
                        id = ?
                        AND LOWER(TRIM(role)) = 'user'
                    LIMIT 1
                    `,
                    [
                        oldUserId
                    ]
                );

            if (
                oldUsers.length === 0
            ) {

                throw new Error(
                    "Old user not found."
                );
            }

            const oldEmployeeId =
                String(
                    oldUsers[0].employee_id
                ).trim();


            await connection.query(
                `
                UPDATE claims
                SET
                    assigned_user_id = ?,
                    user_name = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE
                    TRIM(
                        assigned_user_id
                    )
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
                } catch (e) {
                    console.error(
                        "ROLLBACK ERROR:",
                        e
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

                        id,

                        file_name,

                        uploaded_at,

                        total_claims,

                        CASE
                            WHEN status = 'ACTIVE'
                            THEN 'ACTIVE'
                            ELSE 'DELETED'
                        END AS status

                    FROM upload_batches

                    ORDER BY id DESC
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
                <h2>Upload Dashboard Error</h2>
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

            // =================================================
            // READ WORKBOOK
            // =================================================

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
                workbook.Sheets[
                    sheetName
                ];


            // =================================================
            // SHEET TO JSON
            // =================================================

            const rawRows =
                XLSX.utils.sheet_to_json(
                    sheet,
                    {
                        defval: "",
                        raw: true
                    }
                );


            if (
                rawRows.length === 0
            ) {

                throw new Error(
                    "Excel file is empty."
                );
            }


            // =================================================
            // NORMALIZE EXCEL HEADERS
            // =================================================

            const rows =
                normalizeExcelRows(
                    rawRows
                );


            // =================================================
            // REQUIRED COLUMNS
            // IMPORTANT:
            // Department is checked after normalization
            // =================================================

            const requiredColumns = [

                "CLAIM_REF_NO",

                "INWARD_NO",

                "POLICY_NO",

                "CLAIM_AMT",

                "Vertical",

                "Department",

                "User ID",

                "User Name",

                "Claim Type",

                "Status",

                "Date",

                "Time"
            ];


            const firstRow =
                rows[0];


            const existingColumns =
                Object.keys(
                    firstRow
                );


            const missingColumns =
                requiredColumns.filter(
                    column =>
                        !existingColumns.includes(
                            normalizeHeader(
                                column
                            )
                        )
                );


            if (
                missingColumns.length > 0
            ) {

                throw new Error(
                    "Invalid Excel Format. Missing columns: " +
                    missingColumns.join(", ")
                );
            }


            console.log(
                "EXCEL COLUMNS:",
                existingColumns
            );


            // =================================================
            // DATABASE
            // =================================================

            connection =
                await db.getConnection();

            await connection.beginTransaction();


            // =================================================
            // CREATE UPLOAD BATCH
            // =================================================

            const [batchResult] =
                await connection.query(
                    `
                    INSERT INTO upload_batches
                    (
                        file_name,
                        uploaded_by,
                        total_claims,
                        status,
                        uploaded_at
                    )
                    VALUES
                    (
                        ?,
                        ?,
                        ?,
                        'ACTIVE',
                        NOW()
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
            // INSERT EACH CLAIM
            // =================================================

            for (
                let index = 0;
                index < rows.length;
                index++
            ) {

                const row =
                    rows[index];


                // ---------------------------------------------
                // BASIC
                // ---------------------------------------------

                const claimRefNo =
                    String(
                        excelValue(
                            row,
                            "CLAIM_REF_NO"
                        ) || ""
                    ).trim();


                if (!claimRefNo) {

                    throw new Error(
                        `CLAIM_REF_NO missing at Excel row ${index + 2}.`
                    );
                }


                const inwardNo =
                    String(
                        excelValue(
                            row,
                            "INWARD_NO"
                        ) || ""
                    ).trim();


                const policyNo =
                    String(
                        excelValue(
                            row,
                            "POLICY_NO"
                        ) || ""
                    ).trim();


                const claimAmount =
                    parseExcelNumber(
                        excelValue(
                            row,
                            "CLAIM_AMT"
                        )
                    );


                const vertical =
                    String(
                        excelValue(
                            row,
                            "Vertical"
                        ) || ""
                    ).trim();


                // IMPORTANT:
                // Use actual Department column
                const department =
                    String(
                        excelValue(
                            row,
                            "Department"
                        ) || ""
                    ).trim();


                const employeeId =
                    String(
                        excelValue(
                            row,
                            "User ID"
                        ) || ""
                    ).trim();


                const excelUserName =
                    String(
                        excelValue(
                            row,
                            "User Name"
                        ) || ""
                    ).trim();


                // =================================================
                // USER VALIDATION
                // IMPORTANT:
                // assigned_user_id stores employee_id
                // =================================================

                let assignedUserId =
                    employeeId;

                let userName =
                    excelUserName;


                if (
                    employeeId
                ) {

                    const [userRows] =
                        await connection.query(
                            `
                            SELECT
                                employee_id,
                                username
                            FROM users
                            WHERE
                                TRIM(employee_id)
                                =
                                TRIM(?)

                                AND
                                LOWER(TRIM(role))
                                = 'user'

                                AND
                                is_active = TRUE

                            LIMIT 1
                            `,
                            [
                                employeeId
                            ]
                        );


                    if (
                        userRows.length === 0
                    ) {

                        throw new Error(
                            `Employee ID '${employeeId}' not found or inactive.`
                        );
                    }


                    assignedUserId =
                        String(
                            userRows[0]
                                .employee_id
                        ).trim();

                    userName =
                        userRows[0]
                            .username;
                }


                // =================================================
                // OPTIONAL FIELDS
                // =================================================

                const additionalDeduction =
                    parseExcelNumber(
                        excelValue(
                            row,
                            "Additional Deduction"
                        )
                    );


                const alAmount =
                    parseExcelNumber(
                        excelValue(
                            row,
                            "AL_AMT"
                        )
                    );


                const claimClass =
                    String(
                        excelValue(
                            row,
                            "CLAIM_CLASS"
                        ) || ""
                    ).trim();


                const hospitalCode =
                    String(
                        excelValue(
                            row,
                            "Hospital Code"
                        ) || ""
                    ).trim();


                const typeOfMou =
                    String(
                        excelValue(
                            row,
                            "Type of MOU"
                        ) || ""
                    ).trim();


                const diagnosis =
                    String(
                        excelValue(
                            row,
                            "Diagnosis"
                        ) || ""
                    ).trim();


                const diagnosis2 =
                    String(
                        excelValue(
                            row,
                            "Diagnosis 2"
                        ) || ""
                    ).trim();


                const policyName =
                    String(
                        excelValue(
                            row,
                            "POLICY_NAME"
                        ) || ""
                    ).trim();


                const queue =
                    String(
                        excelValue(
                            row,
                            "Queue"
                        ) || ""
                    ).trim();


                const ageing =
                    String(
                        excelValue(
                            row,
                            "Ageing"
                        ) || ""
                    ).trim();


                // =================================================
                // DATE
                // =================================================

                const rawDate =
                    excelValue(
                        row,
                        "Date"
                    );

                const claimDate =
                    convertExcelDate(
                        rawDate
                    );


                if (
                    rawDate !== "" &&
                    rawDate !== null &&
                    rawDate !== undefined &&
                    !claimDate
                ) {

                    throw new Error(
                        `Invalid Date for Claim ${claimRefNo}.`
                    );
                }


                // =================================================
                // TIME
                // =================================================

                const claimTime =
                    convertExcelTime(
                        excelValue(
                            row,
                            "Time"
                        )
                    );


                const todayStatus =
                    String(
                        excelValue(
                            row,
                            "Today Status"
                        ) || ""
                    ).trim();


                const i3Status =
                    String(
                        excelValue(
                            row,
                            "I3 Status"
                        ) || ""
                    ).trim();


                const fullQc =
                    String(
                        excelValue(
                            row,
                            "Full qc"
                        ) || ""
                    ).trim();


                const relation =
                    String(
                        excelValue(
                            row,
                            "RELATION"
                        ) || ""
                    ).trim();


                const hnf =
                    String(
                        excelValue(
                            row,
                            "HNF"
                        ) || ""
                    ).trim();


                const ilomId =
                    String(
                        excelValue(
                            row,
                            "ILOM ID"
                        ) || ""
                    ).trim() || null;


                const approveAmount =
                    parseExcelNumber(
                        excelValue(
                            row,
                            "Approve AMT"
                        )
                    );


                const remark =
                    String(
                        excelValue(
                            row,
                            "Remark"
                        ) || ""
                    ).trim() || null;


                const deductionAmount =
                    parseExcelNumber(
                        excelValue(
                            row,
                            "Deduction AMT"
                        )
                    );


                const interDocExe =
                    String(
                        excelValue(
                            row,
                            "inter. Doc & Exe"
                        ) || ""
                    ).trim();


                const lot =
                    String(
                        excelValue(
                            row,
                            "lot",
                            "LOT"
                        ) || ""
                    ).trim();


                const platform =
                    String(
                        excelValue(
                            row,
                            "platform",
                            "Platform"
                        ) || ""
                    ).trim();


                // =================================================
                // CLAIM TYPE
                // =================================================

                const claimType =
                    normalizeClaimType(
                        excelValue(
                            row,
                            "Claim Type"
                        )
                    );


                if (
                    claimType !== null &&
                    !VALID_CLAIM_TYPES.includes(
                        claimType
                    )
                ) {

                    throw new Error(
                        `Invalid Claim Type '${claimType}' for Claim ${claimRefNo}.`
                    );
                }


                // =================================================
                // STATUS
                // =================================================

                const claimStatus =
                    normalizeStatus(
                        excelValue(
                            row,
                            "Status"
                        )
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
                // INSERT
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

                        assigned_user_id,
                        user_name,

                        claim_type,
                        ilom_id,

                        approve_amount,
                        claim_status,

                        user_remark,
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

                        ?, ?, ?,

                        ?, ?,

                        ?, ?,

                        ?, ?, ?,

                        ?, ?,

                        ?, ?, ?,

                        ?, ?,

                        ?, ?,

                        ?, ?,

                        ?, ?,

                        ?, ?,

                        ?, ?,

                        ?, ?,

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

                        assignedUserId,
                        userName,

                        claimType,
                        ilomId,

                        approveAmount,
                        claimStatus,

                        remark,
                        deductionAmount,

                        interDocExe,

                        lot,
                        platform
                    ]
                );
            }


            // =================================================
            // COMMIT
            // =================================================

            await connection.commit();


            console.log(
                `EXCEL UPLOAD SUCCESS: ${req.file.originalname}`
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
                WHERE
                    id = ?
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
                await connection.rollback();
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

            return res.status(400).send(
                "Employee ID is missing."
            );
        }


        // =================================================
        // DOUBLE SAVE PROTECTION
        // =================================================

        if (
            savingUsers.has(
                employeeId
            )
        ) {

            return res.status(409).send(`
                <h2>Save Already In Progress</h2>

                <p>
                    Your previous save request is still processing.
                    Please wait.
                </p>

                <a href="/user">
                    Back to Dashboard
                </a>
            `);
        }


        savingUsers.add(
            employeeId
        );


        let connection;

        try {

            // =================================================
            // FIND ONLY SUBMITTED CLAIM IDS
            // =================================================

            const submittedKeys =
                Object.keys(
                    req.body
                );


            const claimIds = [
                ...new Set(

                    submittedKeys

                        .map(
                            key => {

                                const match =
                                    key.match(
                                        /^(?:claim_type|ilom_id|approve_amount|claim_status|user_remark|deduction_amount|diagnosis_2|inter_doc_exe)_(\d+)$/
                                    );

                                return match
                                    ? Number(
                                        match[1]
                                    )
                                    : null;
                            }
                        )

                        .filter(
                            id =>
                                Number.isInteger(
                                    id
                                ) &&
                                id > 0
                        )
                )
            ];


            if (
                claimIds.length === 0
            ) {

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


            connection =
                await db.getConnection();

            await connection.beginTransaction();


            // =================================================
            // UPDATE ONLY SUBMITTED CLAIMS
            // =================================================

            for (
                const id of claimIds
            ) {

                // ---------------------------------------------
                // SECURITY
                // ---------------------------------------------

                const [claimRows] =
                    await connection.query(
                        `
                        SELECT id
                        FROM claims
                        WHERE
                            id = ?

                            AND
                            TRIM(
                                assigned_user_id
                            )
                            =
                            TRIM(?)

                        LIMIT 1
                        `,
                        [
                            id,
                            employeeId
                        ]
                    );


                if (
                    claimRows.length === 0
                ) {

                    throw new Error(
                        `Claim ID ${id} does not belong to this user.`
                    );
                }


                // ---------------------------------------------
                // VALUES
                // ---------------------------------------------

                const claimType =
                    normalizeClaimType(
                        req.body[
                            `claim_type_${id}`
                        ]
                    );


                const ilomId =
                    String(
                        req.body[
                            `ilom_id_${id}`
                        ] || ""
                    ).trim() || null;


                let approveAmount = 0;


                if (
                    req.body[
                        `approve_amount_${id}`
                    ] !== undefined
                ) {

                    approveAmount =
                        parseExcelNumber(
                            req.body[
                                `approve_amount_${id}`
                            ]
                        );
                }


                const claimStatus =
                    normalizeStatus(
                        req.body[
                            `claim_status_${id}`
                        ]
                    );


                if (
                    !VALID_STATUSES.includes(
                        claimStatus
                    )
                ) {

                    throw new Error(
                        `Invalid Status '${claimStatus}' for Claim ID ${id}.`
                    );
                }


                const userRemark =
                    String(
                        req.body[
                            `user_remark_${id}`
                        ] || ""
                    ).trim() || null;


                let deductionAmount = 0;


                if (
                    req.body[
                        `deduction_amount_${id}`
                    ] !== undefined
                ) {

                    deductionAmount =
                        parseExcelNumber(
                            req.body[
                                `deduction_amount_${id}`
                            ]
                        );
                }


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


                // ---------------------------------------------
                // UPDATE
                // ---------------------------------------------

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
                            CURRENT_TIMESTAMP

                    WHERE

                        id = ?

                        AND

                        TRIM(
                            assigned_user_id
                        )
                        =
                        TRIM(?)
                    `,
                    [

                        claimType,

                        ilomId,

                        approveAmount,

                        claimStatus,

                        userRemark,

                        deductionAmount,

                        diagnosis2,

                        interDocExe,

                        id,

                        employeeId
                    ]
                );
            }


            await connection.commit();


            console.log(
                `CLAIMS SAVED SUCCESSFULLY: ${employeeId}`
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
                <h2>Save Claims Failed</h2>

                <pre>${error.message}</pre>

                <br>

                <a href="/user">
                    Back to Dashboard
                </a>
            `);

        } finally {

            if (connection) {
                connection.release();
            }

            savingUsers.delete(
                employeeId
            );
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

                return res.status(400).send(
                    "Employee ID is missing."
                );
            }


            // =================================================
            // CLAIMS
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
                        ON c.upload_batch_id = ub.id

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
            // USER SUMMARY
            // =================================================

            const [[userSummary]] =
                await db.query(
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
                                'Investigation&Query'
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
                                'SentBack'
                            ),
                            0
                        ) AS sent_back,

                        COALESCE(
                            SUM(
                                claim_status =
                                'Keep'
                            ),
                            0
                        ) AS keep_count,

                        COALESCE(
                            SUM(
                                claim_status =
                                'OtherDoctor/Executive'
                            ),
                            0
                        ) AS other_doctor_executive,

                        COALESCE(
                            SUM(
                                claim_status =
                                'ROD Cancel'
                            ),
                            0
                        ) AS rod_cancel,

                        COALESCE(
                            SUM(
                                claim_status <>
                                'Pending'
                            ),
                            0
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
            // PLATFORM SUMMARY
            // =================================================

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
                                claim_status =
                                'Approved'
                            ),
                            0
                        ) AS approved,

                        COALESCE(
                            SUM(
                                claim_status =
                                'Rejected'
                            ),
                            0
                        ) AS rejected,

                        COALESCE(
                            SUM(
                                claim_status =
                                'Query'
                            ),
                            0
                        ) AS query_count,

                        COALESCE(
                            SUM(
                                claim_status =
                                'Re-Query'
                            ),
                            0
                        ) AS requery,

                        COALESCE(
                            SUM(
                                claim_status =
                                'Investigation&Query'
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
                                'SentBack'
                            ),
                            0
                        ) AS sent_back,

                        COALESCE(
                            SUM(
                                claim_status =
                                'Keep'
                            ),
                            0
                        ) AS keep_count,

                        COALESCE(
                            SUM(
                                claim_status =
                                'OtherDoctor/Executive'
                            ),
                            0
                        ) AS other_doctor_executive,

                        COALESCE(
                            SUM(
                                claim_status =
                                'ROD Cancel'
                            ),
                            0
                        ) AS rod_cancel,

                        COALESCE(
                            SUM(
                                claim_status =
                                'Pending'
                            ),
                            0
                        ) AS pending,

                        COALESCE(
                            SUM(
                                claim_status <>
                                'Pending'
                            ),
                            0
                        ) AS total_productivity

                    FROM claims

                    WHERE
                        TRIM(
                            assigned_user_id
                        )
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


            // =================================================
            // FORMAT
            // =================================================

            const formattedClaims =
                claims.map(
                    claim => {

                        let formattedDate =
                            "-";

                        let formattedTime =
                            "-";

                        let formattedUploadedAt =
                            "-";


                        if (
                            claim.claim_date
                        ) {

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


                        if (
                            claim.claim_time
                        ) {

                            const time =
                                String(
                                    claim.claim_time
                                ).trim();

                            const match =
                                time.match(
                                    /^(\d{2}):(\d{2})(?::(\d{2}))?$/
                                );

                            if (match) {

                                let hour =
                                    Number(
                                        match[1]
                                    );

                                const minute =
                                    match[2];

                                const suffix =
                                    hour >= 12
                                        ? "PM"
                                        : "AM";

                                hour =
                                    hour % 12 || 12;

                                formattedTime =
                                    `${String(hour).padStart(2, "0")}:${minute} ${suffix}`;

                            } else {

                                formattedTime =
                                    time;
                            }
                        }


                        if (
                            claim.uploaded_at
                        ) {

                            const date =
                                new Date(
                                    claim.uploaded_at
                                );

                            if (
                                !isNaN(
                                    date.getTime()
                                )
                            ) {

                                formattedUploadedAt =
                                    date.toLocaleString(
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
                    }
                );


            // =================================================
            // SUMMARY
            // =================================================

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
                    )
            };


            let savedAt =
                null;


            if (
                req.query.saved === "1"
            ) {

                savedAt =
                    new Date()
                        .toLocaleString(
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
                        processSummary,

                    platformSummary:
                        platformSummary
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
            normalizeRole(
                req.session.user.role
            ) !== "admin"
        ) {

            return res.redirect("/");
        }


        const fromDate =
            String(
                req.query.fromDate || ""
            ).trim();


        const toDate =
            String(
                req.query.toDate || ""
            ).trim();


        try {

            let sql = `
                SELECT

                    claim_ref_no AS CLAIM_REF_NO,

                    inward_no AS INWARD_NO,

                    policy_no AS POLICY_NO,

                    claim_amount AS CLAIM_AMT,

                    vertical AS Vertical,

                    department AS Department,

                    additional_deduction
                        AS "Additional Deduction",

                    al_amount AS AL_AMT,

                    claim_class AS CLAIM_CLASS,

                    hospital_code
                        AS "Hospital Code",

                    type_of_mou
                        AS "Type of MOU",

                    diagnosis AS Diagnosis,

                    diagnosis_2
                        AS "Diagnosis 2",

                    policy_name AS POLICY_NAME,

                    queue AS Queue,

                    ageing AS Ageing,

                    claim_date AS Date,

                    claim_time AS Time,

                    today_status
                        AS "Today Status",

                    i3_status
                        AS "I3 Status",

                    full_qc
                        AS "Full qc",

                    relation AS RELATION,

                    hnf AS HNF,

                    assigned_user_id
                        AS "User ID",

                    user_name
                        AS "User Name",

                    claim_type
                        AS "Claim Type",

                    ilom_id
                        AS "ILOM ID",

                    approve_amount
                        AS "Approve AMT",

                    claim_status
                        AS Status,

                    user_remark
                        AS Remark,

                    deduction_amount
                        AS "Deduction AMT",

                    inter_doc_exe
                        AS "inter. Doc & Exe",

                    lot AS lot,

                    platform AS platform

                FROM claims
            `;


            const params = [];


            // =================================================
            // DATE FILTER ONLY IF PROVIDED
            // =================================================

            if (
                fromDate &&
                toDate
            ) {

                sql += `
                    WHERE
                        DATE(
                            COALESCE(
                                claim_date,
                                created_at
                            )
                        )
                        BETWEEN ?
                        AND ?
                `;

                params.push(
                    fromDate,
                    toDate
                );
            }


            sql += `
                ORDER BY id DESC
            `;


            const [claims] =
                await db.query(
                    sql,
                    params
                );


            console.log(
                "ADMIN DOWNLOAD ROW COUNT:",
                claims.length
            );


            const worksheet =
                XLSX.utils.json_to_sheet(
                    claims
                );


            worksheet["!cols"] = [

                { wch: 20 },
                { wch: 18 },
                { wch: 18 },
                { wch: 15 },
                { wch: 15 },
                { wch: 18 },
                { wch: 18 },
                { wch: 15 },
                { wch: 18 },
                { wch: 20 },
                { wch: 18 },
                { wch: 20 },
                { wch: 20 },
                { wch: 20 },
                { wch: 15 },
                { wch: 15 },
                { wch: 15 },
                { wch: 18 },
                { wch: 15 },
                { wch: 20 },
                { wch: 20 },
                { wch: 15 },
                { wch: 15 },
                { wch: 18 },
                { wch: 20 },
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
                "attachment; filename=updated-claims.xlsx"
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
            () => {

                res.redirect("/");
            }
        );
    }
);


// =====================================================
// START SERVER
// =====================================================

app.listen(
    PORT,
    async () => {

        console.log(
            `Server running on port ${PORT}`
        );

        await testDatabase();
    }
);