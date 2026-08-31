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
// EXPRESS SETTINGS
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

        secret: "mdi-claim-secret-key",

        resave: false,

        saveUninitialized: false,

        cookie: {
            maxAge: 1000 * 60 * 60
        }
    })
);

// =====================================================
// CONSTANTS
// IMPORTANT: THESE MATCH YOUR MYSQL ENUM VALUES
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
// TEST DATABASE
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

        console.log(
            "MySQL Connection Failed"
        );

        console.log(
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

    const {
        username,
        password
    } = req.body;

    try {

        const [users] = await db.query(
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

        if (users.length === 0) {

            return res.render(
                "login",
                {
                    error:
                        "Invalid username or password"
                }
            );
        }

        const user = users[0];

        // =================================================
        // SAVE LOGIN USER IN SESSION
        // =================================================

        req.session.user = {

            id: user.id,

            employee_id: user.employee_id,

            username: user.username,

            role: user.role
        };

        console.log(
            "LOGIN SUCCESS:",
            req.session.user
        );

        // =================================================
        // REDIRECT BASED ON ROLE
        // =================================================

        if (user.role === "admin") {

            return res.redirect("/admin");
        }

        if (user.role === "upload") {

            return res.redirect("/upload");
        }

        if (user.role === "user") {

            return res.redirect("/user");
        }

        return res.redirect("/");

    } catch (error) {

        console.error(
            "LOGIN ERROR:",
            error
        );

        res.render(
            "login",
            {
                error:
                    "Server error. Please try again."
            }
        );
    }
});

// =====================================================
// ADMIN DASHBOARD
// =====================================================

app.get("/admin", async (req, res) => {

    if (
        !req.session.user ||
        req.session.user.role !== "admin"
    ) {

        return res.redirect("/");
    }

    try {

        // =================================================
        // TOP SUMMARY COUNTS
        // =================================================

        const [[total]] = await db.query(`
            SELECT COUNT(*) AS count
            FROM claims
        `);

        const [[pending]] = await db.query(`
            SELECT COUNT(*) AS count
            FROM claims
            WHERE claim_status = 'Pending'
        `);

        const [[approved]] = await db.query(`
            SELECT COUNT(*) AS count
            FROM claims
            WHERE claim_status = 'Approved'
        `);

        const [[rejected]] = await db.query(`
            SELECT COUNT(*) AS count
            FROM claims
            WHERE claim_status = 'Rejected'
        `);

        const [[query]] = await db.query(`
            SELECT COUNT(*) AS count
            FROM claims
            WHERE claim_status = 'Query'
        `);

        const [[requery]] = await db.query(`
            SELECT COUNT(*) AS count
            FROM claims
            WHERE claim_status = 'Re-Query'
        `);

        const [[investigationQuery]] = await db.query(`
            SELECT COUNT(*) AS count
            FROM claims
            WHERE claim_status = 'Query & Investigation'
        `);

        const [[investigation]] = await db.query(`
            SELECT COUNT(*) AS count
            FROM claims
            WHERE claim_status = 'Investigation'
        `);

        const [[sentBack]] = await db.query(`
            SELECT COUNT(*) AS count
            FROM claims
            WHERE claim_status = 'Sent-Back'
        `);

        const [[keep]] = await db.query(`
            SELECT COUNT(*) AS count
            FROM claims
            WHERE claim_status = 'Keep'
        `);

        const [[otherDoctorExecutive]] = await db.query(`
            SELECT COUNT(*) AS count
            FROM claims
            WHERE claim_status = 'Other-Doctor/Executive'
        `);

        const [[rodCancel]] = await db.query(`
            SELECT COUNT(*) AS count
            FROM claims
            WHERE claim_status = 'ROD-Cancel'
        `);

        // =================================================
        // TOTAL USERS
        // =================================================

        const [[users]] = await db.query(`
            SELECT COUNT(*) AS count
            FROM users
            WHERE role = 'user'
        `);

        // =================================================
        // USER LIST
        // =================================================

        const [userList] = await db.query(`
            SELECT
                id,
                employee_id,
                username,
                department,
                is_active
            FROM users
            WHERE role = 'user'
            ORDER BY username
        `);

        // =================================================
        // PRODUCTIVITY SUMMARY
        // =================================================

        const [processSummary] = await db.query(`
            SELECT

                COALESCE(c.platform, '-') AS platform,

                COALESCE(
                    u.username,
                    c.user_name,
                    '-'
                ) AS user_name,

                COUNT(*) AS total_allocated,

                SUM(
                    CASE
                        WHEN c.claim_status = 'Approved'
                        THEN 1
                        ELSE 0
                    END
                ) AS approved,

                SUM(
                    CASE
                        WHEN c.claim_status = 'Rejected'
                        THEN 1
                        ELSE 0
                    END
                ) AS rejected,

                SUM(
                    CASE
                        WHEN c.claim_status = 'Query'
                        THEN 1
                        ELSE 0
                    END
                ) AS query_count,

                SUM(
                    CASE
                        WHEN c.claim_status = 'Re-Query'
                        THEN 1
                        ELSE 0
                    END
                ) AS requery,

                SUM(
                    CASE
                        WHEN c.claim_status =
                            'Query & Investigation'
                        THEN 1
                        ELSE 0
                    END
                ) AS investigation_query,

                SUM(
                    CASE
                        WHEN c.claim_status IN (
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
                        THEN 1
                        ELSE 0
                    END
                ) AS total_productivity,

                SUM(
                    CASE
                        WHEN c.claim_status = 'Investigation'
                        THEN 1
                        ELSE 0
                    END
                ) AS investigation,

                SUM(
                    CASE
                        WHEN c.claim_status = 'Sent-Back'
                        THEN 1
                        ELSE 0
                    END
                ) AS sent_back,

                SUM(
                    CASE
                        WHEN c.claim_status = 'Keep'
                        THEN 1
                        ELSE 0
                    END
                ) AS keep_count,

                SUM(
                    CASE
                        WHEN c.claim_status =
                            'Other-Doctor/Executive'
                        THEN 1
                        ELSE 0
                    END
                ) AS other_doctor_executive,

                SUM(
                    CASE
                        WHEN c.claim_status = 'ROD-Cancel'
                        THEN 1
                        ELSE 0
                    END
                ) AS rod_cancel,

                SUM(
                    CASE
                        WHEN c.claim_status = 'Pending'
                        THEN 1
                        ELSE 0
                    END
                ) AS pending

            FROM claims c

            LEFT JOIN users u
                ON c.assigned_user_id = u.employee_id

            GROUP BY
                c.platform,
                COALESCE(
                    u.username,
                    c.user_name,
                    '-'
                )

            ORDER BY
                c.platform,
                user_name
        `);

        // =================================================
        // RENDER ADMIN
        // =================================================

        res.render(
            "admin-dashboard",
            {

                user:
                    req.session.user,

                counts: {

                    total:
                        total.count,

                    pending:
                        pending.count,

                    approved:
                        approved.count,

                    rejected:
                        rejected.count,

                    query:
                        query.count,

                    requery:
                        requery.count,

                    investigationQuery:
                        investigationQuery.count,

                    investigation:
                        investigation.count,

                    sentBack:
                        sentBack.count,

                    keep:
                        keep.count,

                    otherDoctorExecutive:
                        otherDoctorExecutive.count,

                    rodCancel:
                        rodCancel.count,

                    users:
                        users.count
                },

                userList:
                    userList,

                processSummary:
                    processSummary
            }
        );

    } catch (error) {

        console.error(
            "ADMIN ERROR:",
            error
        );

        res.status(500).send(`

            <h2>Database Error</h2>

            <pre>${error.message}</pre>

            <br>

            <a href="/admin">
                Back to Admin
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
            req.session.user.role !== "admin"
        ) {

            return res.redirect("/");
        }

        const {
            username,
            password,
            department,
            employee_id
        } = req.body;

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
                    employee_id || null,
                    username,
                    password,
                    department
                ]
            );

            res.redirect("/admin");

        } catch (error) {

            console.error(
                "CREATE USER ERROR:",
                error
            );

            res.status(500).send(`

                <h2>Failed to create user</h2>

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

app.post("/admin/reassign", async (req, res) => {

    if (
        !req.session.user ||
        req.session.user.role !== "admin"
    ) {
        return res.redirect("/");
    }

    const { oldUserId, newUserId } = req.body;

    if (!oldUserId || !newUserId) {
        return res.status(400).send(`
            <h2>Invalid User Selection</h2>
            <p>Please select both users.</p>
            <a href="/admin">Back to Admin</a>
        `);
    }

    if (String(oldUserId) === String(newUserId)) {
        return res.status(400).send(`
            <h2>Invalid Reassignment</h2>
            <p>Both users cannot be the same.</p>
            <a href="/admin">Back to Admin</a>
        `);
    }

    let connection;

    try {

        connection = await db.getConnection();

        await connection.beginTransaction();

        // ---------------------------------------------
        // OLD USER - GET EMPLOYEE ID
        // ---------------------------------------------

        const [oldUser] = await connection.query(
            `
            SELECT employee_id
            FROM users
            WHERE id = ?
            AND role = 'user'
            LIMIT 1
            `,
            [oldUserId]
        );

        if (oldUser.length === 0) {
            throw new Error("Old user not found.");
        }

        const oldEmployeeId = oldUser[0].employee_id;

        // ---------------------------------------------
        // NEW USER - GET EMPLOYEE ID
        // ---------------------------------------------

        const [newUser] = await connection.query(
            `
            SELECT employee_id
            FROM users
            WHERE id = ?
            AND role = 'user'
            AND is_active = TRUE
            LIMIT 1
            `,
            [newUserId]
        );

        if (newUser.length === 0) {
            throw new Error("New user not found or inactive.");
        }

        const newEmployeeId = newUser[0].employee_id;

        // ---------------------------------------------
        // REASSIGN CLAIMS
        // ---------------------------------------------

        const [result] = await connection.query(
            `
            UPDATE claims
            SET
                assigned_user_id = ?,
                user_name = (
                    SELECT username
                    FROM users
                    WHERE employee_id = ?
                    LIMIT 1
                ),
                updated_at = CURRENT_TIMESTAMP
            WHERE assigned_user_id = ?
            `,
            [
                newEmployeeId,
                newEmployeeId,
                oldEmployeeId
            ]
        );

        await connection.commit();

        console.log("=================================");
        console.log("REASSIGN SUCCESS");
        console.log("Old Employee ID:", oldEmployeeId);
        console.log("New Employee ID:", newEmployeeId);
        console.log("Claims Updated:", result.affectedRows);
        console.log("=================================");

        res.redirect("/admin");

    } catch (error) {

        if (connection) {
            await connection.rollback();
        }

        console.error("REASSIGN ERROR:", error);

        res.status(500).send(`
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
});
// =====================================================
// UPLOAD DASHBOARD
// =====================================================

app.get(
    "/upload",
    async (req, res) => {

        if (
            !req.session.user ||
            req.session.user.role !== "upload"
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

            res.render(
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

            res.status(500).send(`

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
            req.session.user.role !== "upload"
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
                        type: "buffer"
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
                        !excelColumns.includes(column)
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
                    String(
                        row["Vertical"] || ""
                    ).trim();

                // =================================================
                // EMPLOYEE ID
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
                            AND role = 'user'
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

                const claimDate =
                    String(
                        row["Date"] || ""
                    ).trim();

                const claimTime =
                    String(
                        row["Time"] || ""
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

                // =================================================
                // CLAIM TYPE
                // =================================================

                let claimType =
                    String(
                        row["Claim Type"] || ""
                    ).trim();

                if (claimType === "") {

                    claimType = null;

                } else {

                    const upperClaimType =
                        claimType.toUpperCase();

                    if (
                        upperClaimType ===
                        "INPATIENT"
                    ) {

                        claimType = "IPD";

                    } else if (
                        upperClaimType ===
                        "OUTPATIENT"
                    ) {

                        claimType = "OPD";

                    } else if (
                        upperClaimType ===
                        "PREPOST"
                    ) {

                        claimType = "Pre Post";

                    } else if (
                        upperClaimType ===
                        "PRE POST"
                    ) {

                        claimType = "Pre Post";

                    } else if (
                        !VALID_CLAIM_TYPES.includes(
                            claimType
                        )
                    ) {

                        throw new Error(
                            `Invalid Claim Type '${claimType}' for Claim ${claimRefNo}.`
                        );
                    }
                }

                // =================================================
                // STATUS
                // =================================================

                let claimStatus =
                    String(
                        row["Status"] || ""
                    ).trim();

                if (claimStatus === "") {

                    claimStatus = "Pending";

                } else {

                    // Normalize Excel values
                    // TO MATCH MYSQL ENUM

                    if (
                        claimStatus === "SentBack"
                    ) {

                        claimStatus = "Sent-Back";
                    }

                    if (
                        claimStatus === "ROD Cancel"
                    ) {

                        claimStatus = "ROD-Cancel";
                    }

                    if (
                        claimStatus ===
                        "Investigation&Query"
                    ) {

                        claimStatus =
                            "Query & Investigation";
                    }

                    if (
                        claimStatus ===
                        "OtherDoctor/Executive"
                    ) {

                        claimStatus =
                            "Other-Doctor/Executive";
                    }
                }

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
                // IMPORTANT:
                // claims TABLE HAS user_name, NOT username
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

            // =================================================
            // COMMIT
            // =================================================

            await connection.commit();

            console.log(
                "EXCEL UPLOAD SUCCESS"
            );

            res.redirect("/upload");

        } catch (error) {

            if (connection) {

                await connection.rollback();
            }

            console.error(
                "EXCEL UPLOAD ERROR:",
                error
            );

            res.status(500).send(`

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
            req.session.user.role !== "upload"
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

            res.redirect("/upload");

        } catch (error) {

            if (connection) {

                await connection.rollback();
            }

            console.error(
                "DELETE UPLOAD ERROR:",
                error
            );

            res.status(500).send(`

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
// SAVE USER CLAIMS
// =====================================================

app.post(
    "/save-claims",
    async (req, res) => {

        if (!req.session.user) {

            return res.redirect("/");
        }

        // =================================================
        // IMPORTANT
        // claims.assigned_user_id contains employee_id
        // =================================================

        const employeeId =
            req.session.user.employee_id;

        if (!employeeId) {

            return res.status(400).send(`

                <h2>Save Failed</h2>

                <p>
                    Employee ID is missing from login session.
                </p>

                <a href="/user">
                    Back to Dashboard
                </a>

            `);
        }

        let connection;

        try {

            console.log(
                "SAVE REQUEST EMPLOYEE ID:",
                employeeId
            );

            console.log(
                "SAVE BODY:",
                req.body
            );

            connection =
                await db.getConnection();

            await connection.beginTransaction();

            // =================================================
            // GET ONLY THIS USER'S CLAIMS
            // =================================================

            const [claims] =
                await connection.query(
                    `
                    SELECT id
                    FROM claims
                    WHERE assigned_user_id = ?
                    `,
                    [
                        employeeId
                    ]
                );

            if (
                !claims ||
                claims.length === 0
            ) {

                await connection.rollback();

                return res.status(400).send(`

                    <h2>Save Failed</h2>

                    <p>
                        No claims found for employee ID:
                        ${employeeId}
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

                // =================================================
                // GET FORM VALUES
                // =================================================

                const claimType =
                    req.body[
                        `claim_type_${id}`
                    ] || null;

                const ilomId =
                    req.body[
                        `ilom_id_${id}`
                    ] || null;

                const approveAmount =
                    req.body[
                        `approve_amount_${id}`
                    ];

                const claimStatus =
                    req.body[
                        `claim_status_${id}`
                    ] || "Pending";

                const userRemark =
                    req.body[
                        `user_remark_${id}`
                    ] || null;

                const deductionAmount =
                    req.body[
                        `deduction_amount_${id}`
                    ];

                const diagnosis2 =
                    req.body[
                        `diagnosis_2_${id}`
                    ] || null;

                const interDocExe =
                    req.body[
                        `inter_doc_exe_${id}`
                    ] || null;

                // =================================================
                // VALIDATE CLAIM TYPE
                // =================================================

                let finalClaimType =
                    claimType;

                if (
                    finalClaimType === ""
                ) {

                    finalClaimType = null;
                }

                if (
                    finalClaimType &&
                    !VALID_CLAIM_TYPES.includes(
                        finalClaimType
                    )
                ) {

                    throw new Error(
                        `Invalid Claim Type for Claim ID ${id}`
                    );
                }

                // =================================================
                // VALIDATE STATUS
                // =================================================

                let finalClaimStatus =
                    claimStatus;

                // Normalize possible frontend values

                if (
                    finalClaimStatus === "SentBack"
                ) {

                    finalClaimStatus =
                        "Sent-Back";
                }

                if (
                    finalClaimStatus ===
                    "ROD Cancel"
                ) {

                    finalClaimStatus =
                        "ROD-Cancel";
                }

                if (
                    finalClaimStatus ===
                    "Investigation&Query"
                ) {

                    finalClaimStatus =
                        "Query & Investigation";
                }

                if (
                    finalClaimStatus ===
                    "OtherDoctor/Executive"
                ) {

                    finalClaimStatus =
                        "Other-Doctor/Executive";
                }

                if (
                    !VALID_STATUSES.includes(
                        finalClaimStatus
                    )
                ) {

                    throw new Error(
                        `Invalid Status '${finalClaimStatus}' for Claim ID ${id}`
                    );
                }

                // =================================================
                // DECIMAL VALUES
                // =================================================

                const finalApproveAmount =
                    approveAmount === "" ||
                    approveAmount === undefined ||
                    approveAmount === null
                        ? 0
                        : Number(approveAmount);

                const finalDeductionAmount =
                    deductionAmount === "" ||
                    deductionAmount === undefined ||
                    deductionAmount === null
                        ? 0
                        : Number(deductionAmount);

                // =================================================
                // VALIDATE NUMBERS
                // =================================================

                if (
                    !Number.isFinite(
                        finalApproveAmount
                    )
                ) {

                    throw new Error(
                        `Invalid Approve Amount for Claim ID ${id}`
                    );
                }

                if (
                    !Number.isFinite(
                        finalDeductionAmount
                    )
                ) {

                    throw new Error(
                        `Invalid Deduction Amount for Claim ID ${id}`
                    );
                }

                // =================================================
                // UPDATE CLAIM
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
                        AND assigned_user_id = ?
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
                `Claims saved successfully for employee ${employeeId}`
            );

            // =================================================
            // REDIRECT
            // =================================================

            return res.redirect(
                "/user?saved=1"
            );

        } catch (error) {

            console.error(
                "SAVE CLAIMS ERROR:",
                error
            );

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
// ADMIN DOWNLOAD UPDATED CLAIMS
// =====================================================

app.get(
    "/admin/download-claims",
    async (req, res) => {

        if (
            !req.session.user ||
            req.session.user.role !== "admin"
        ) {

            return res.redirect("/");
        }

        try {

            const [claims] =
                await db.query(`
                    SELECT

                        claim_ref_no
                            AS CLAIM_REF_NO,

                        inward_no
                            AS INWARD_NO,

                        policy_no
                            AS POLICY_NO,

                        vertical
                            AS Vertical,

                        additional_deduction
                            AS "Additional Deduction",

                        claim_amount
                            AS CLAIM_AMT,

                        al_amount
                            AS AL_AMT,

                        claim_class
                            AS CLAIM_CLASS,

                        hospital_code
                            AS "Hospital Code",

                        type_of_mou
                            AS "Type of MOU",

                        diagnosis
                            AS Diagnosis,

                        diagnosis_2
                            AS "Diagnosis 2",

                        policy_name
                            AS POLICY_NAME,

                        queue
                            AS Queue,

                        ageing
                            AS Ageing,

                        claim_date
                            AS Date,

                        claim_time
                            AS Time,

                        today_status
                            AS "Today Status",

                        i3_status
                            AS "I3 Status",

                        full_qc
                            AS "Full qc",

                        relation
                            AS RELATION,

                        hnf
                            AS HNF,

                        assigned_user_id
                            AS "User ID",

                        lot
                            AS lot,

                        platform
                            AS platform,

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
                            AS "inter. Doc & Exe"

                    FROM claims

                    ORDER BY id DESC
                `);

            // =================================================
            // CREATE EXCEL
            // =================================================

            const worksheet =
                XLSX.utils.json_to_sheet(
                    claims
                );

            const workbook =
                XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(
                workbook,
                worksheet,
                "Claims"
            );

            const excelBuffer =
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

            res.send(
                excelBuffer
            );

        } catch (error) {

            console.error(
                "DOWNLOAD ERROR:",
                error
            );

            res.status(500).send(`

                <h2>Download Failed</h2>

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
// USER DASHBOARD
// =====================================================

app.get(
    "/user",
    async (req, res) => {

        if (!req.session.user) {

            return res.redirect("/");
        }

        try {

            // =================================================
            // IMPORTANT:
            // assigned_user_id = employee_id
            // =================================================

            const employeeId =
                req.session.user.employee_id;

            if (!employeeId) {

                return res.status(400).send(`

                    <h2>User Dashboard Error</h2>

                    <p>
                        Employee ID is missing from login session.
                    </p>

                    <a href="/">
                        Login Again
                    </a>

                `);
            }

            // =================================================
            // GET USER CLAIMS
            // ONLY TODAY'S UPLOADS
            // =================================================

            const [claims] =
                await db.query(
                    `
                    SELECT
                        c.*,

                        u.id AS employeeid,

                        u.employee_id AS employee_id,

                        u.username AS employee_name

                    FROM claims c

                    LEFT JOIN users u
                        ON c.assigned_user_id =
                           u.employee_id

                    WHERE
                        c.assigned_user_id = ?

                        AND c.uploaded_at >= CURDATE()

                    ORDER BY c.id DESC
                    `,
                    [
                        employeeId
                    ]
                );

            // =================================================
            // FORMAT EXCEL DATE / TIME
            // =================================================

            const formattedClaims =
                claims.map(
                    claim => {

                        let formattedDate =
                            "-";

                        let formattedTime =
                            "-";

                        // =================================================
                        // DATE
                        // =================================================

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
                                            day:
                                                "2-digit",

                                            month:
                                                "2-digit",

                                            year:
                                                "numeric"
                                        }
                                    );
                            }
                        }

                        // =================================================
                        // TIME
                        // =================================================

                        if (
                            claim.claim_time
                        ) {

                            let timeValue =
                                String(
                                    claim.claim_time
                                );

                            // If MySQL returns HH:MM:SS
                            if (
                                /^\d{2}:\d{2}:\d{2}$/.test(
                                    timeValue
                                )
                            ) {

                                const time =
                                    new Date(
                                        `1970-01-01T${timeValue}`
                                    );

                                if (
                                    !isNaN(
                                        time.getTime()
                                    )
                                ) {

                                    formattedTime =
                                        time.toLocaleTimeString(
                                            "en-IN",
                                            {
                                                hour:
                                                    "2-digit",

                                                minute:
                                                    "2-digit",

                                                hour12:
                                                    true
                                            }
                                        );
                                }

                            } else {

                                formattedTime =
                                    timeValue;
                            }
                        }

                        return {

                            ...claim,

                            formatted_claim_date:
                                formattedDate,

                            formatted_claim_time:
                                formattedTime
                        };
                    }
                );

            // =================================================
            // DEBUG
            // =================================================

            console.log(
                "USER EMPLOYEE ID:",
                employeeId
            );

            console.log(
                "CLAIMS:",
                formattedClaims.length
            );

            // =================================================
            // RENDER
            // =================================================

            res.render(
                "user-dashboard",
                {

                    user:
                        req.session.user,

                    claims:
                        formattedClaims,

                    saved:
                        req.query.saved === "1"
                }
            );

        } catch (error) {

            console.error(
                "USER DASHBOARD ERROR:",
                error
            );

            res.status(500).send(`

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

app.listen(PORT, "0.0.0.0", async () => {
    console.log(`Server running on port ${PORT}`);
    await testDatabase();
});