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
// DATABASE TEST
// =====================================================

async function testDatabase() {

    try {

        console.log("========== MYSQL DEBUG ==========");

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

        console.log("=================================");

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

    console.log("=================================");
    console.log("LOGIN ATTEMPT");
    console.log("Employee ID:", employee_id);
    console.log("=================================");

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

        console.log(
            "USER FOUND:",
            users
        );

        if (users.length === 0) {

            console.log(
                "LOGIN FAILED - EMPLOYEE ID NOT FOUND:",
                employee_id
            );

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

            console.log(
                "LOGIN FAILED - PASSWORD MISMATCH:",
                employee_id
            );

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

        console.log(
            "LOGIN SUCCESS:",
            req.session.user
        );

        if (role === "admin") {

            return res.redirect(
                "/admin"
            );
        }

        if (role === "upload") {

            return res.redirect(
                "/upload"
            );
        }

        if (role === "user") {

            return res.redirect(
                "/user"
            );
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
// ADMIN DASHBOARD
// =====================================================

app.get("/admin", async (req, res) => {

    if (
        !req.session.user ||
        normalizeRole(req.session.user.role) !== "admin"
    ) {
        return res.redirect("/");
    }

    try {

        // =================================================
        // OVERALL CLAIM COUNTS
        // =================================================

        const [[summary]] = await db.query(`
            SELECT

                COUNT(*) AS total,

                SUM(claim_status = 'Pending') AS pending,

                SUM(claim_status = 'Approved') AS approved,

                SUM(claim_status = 'Rejected') AS rejected,

                SUM(claim_status = 'Query') AS query_count,

                SUM(claim_status = 'Re-Query') AS requery,

                SUM(
                    claim_status = 'Query & Investigation'
                ) AS investigation_query,

                SUM(
                    claim_status = 'Investigation'
                ) AS investigation,

                SUM(
                    claim_status = 'Sent-Back'
                ) AS sent_back,

                SUM(
                    claim_status = 'Keep'
                ) AS keep_count,

                SUM(
                    claim_status = 'Other-Doctor/Executive'
                ) AS other_doctor_executive,

                SUM(
                    claim_status = 'ROD-Cancel'
                ) AS rod_cancel,

                SUM(
                    claim_status <> 'Pending'
                ) AS processed

            FROM claims
        `);

        // =================================================
        // TOTAL ACTIVE USERS
        // =================================================

        const [[userCount]] = await db.query(`
            SELECT COUNT(*) AS count
            FROM users
            WHERE LOWER(TRIM(role)) = 'user'
            AND is_active = TRUE
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
            WHERE LOWER(TRIM(role)) = 'user'
            ORDER BY username
        `);

        // =================================================
        // PROCESS SUMMARY
        // =================================================

        const [processSummary] = await db.query(`
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
                ON TRIM(c.assigned_user_id)
                = TRIM(u.employee_id)

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

        // =================================================
        // ADMIN USER OBJECT
        // =================================================

        const adminUser = {
            ...req.session.user,
            name: req.session.user.username
        };

        // =================================================
        // RENDER
        // =================================================

        return res.render(
            "admin-dashboard",
            {

                user: adminUser,

                counts: {

                    total:
                        Number(summary.total || 0),

                    pending:
                        Number(summary.pending || 0),

                    approved:
                        Number(summary.approved || 0),

                    rejected:
                        Number(summary.rejected || 0),

                    query:
                        Number(summary.query_count || 0),

                    requery:
                        Number(summary.requery || 0),

                    investigationQuery:
                        Number(
                            summary.investigation_query || 0
                        ),

                    investigation:
                        Number(
                            summary.investigation || 0
                        ),

                    sentBack:
                        Number(summary.sent_back || 0),

                    keep:
                        Number(summary.keep_count || 0),

                    otherDoctorExecutive:
                        Number(
                            summary.other_doctor_executive || 0
                        ),

                    rodCancel:
                        Number(summary.rod_cancel || 0),

                    processed:
                        Number(summary.processed || 0),

                    users:
                        Number(userCount.count || 0)
                },

                userList: userList,

                processSummary: processSummary
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
});
// =========================================================
// DOWNLOAD PROCESS SUMMARY
// =========================================================

app.get("/admin/download-process-summary", async (req, res) => {

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

                SUM(c.claim_status = 'Pending') AS pending,

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
                ON TRIM(c.assigned_user_id)
                = TRIM(u.employee_id)

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

        // Excel workbook
        const workbook = XLSX.utils.book_new();

        const excelData = rows.map(row => ({
            "Platform": row.platform || "-",
            "Employee ID": row.employee_id || "-",
            "User Name": row.user_name || "-",
            "Total Allocated": Number(row.total_allocated || 0),
            "Approved": Number(row.approved || 0),
            "Rejected": Number(row.rejected || 0),
            "Query": Number(row.query_count || 0),
            "Re-Query": Number(row.requery || 0),
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
            "ROD Cancel":
                Number(row.rod_cancel || 0),
            "Pending":
                Number(row.pending || 0)
        }));

        const worksheet =
            XLSX.utils.json_to_sheet(excelData);

        // Column widths
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

        // Convert to Excel buffer
        const buffer = XLSX.write(
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
            <h2>Process Summary Download Error</h2>
            <pre>${error.message}</pre>
            <br>
            <a href="/admin">Back to Admin</a>
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
            normalizeRole(req.session.user.role) !== "admin"
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

            // =================================================
            // CHECK EMPLOYEE ID
            // =================================================

            const [employeeExists] =
                await db.query(`
                    SELECT id
                    FROM users
                    WHERE LOWER(TRIM(employee_id))
                        = LOWER(TRIM(?))
                    LIMIT 1
                `, [
                    employeeId
                ]);

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

            // =================================================
            // CHECK USERNAME
            // =================================================

            const [usernameExists] =
                await db.query(`
                    SELECT id
                    FROM users
                    WHERE LOWER(TRIM(username))
                        = LOWER(TRIM(?))
                    LIMIT 1
                `, [
                    username
                ]);

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

            // =================================================
            // INSERT USER
            // =================================================

            await db.query(`
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
            `, [
                employeeId,
                username,
                password,
                department || null
            ]);

            return res.redirect("/admin");

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
            normalizeRole(req.session.user.role) !== "admin"
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

            // =================================================
            // OLD USER
            // =================================================

            const [oldUser] =
                await connection.query(`
                    SELECT employee_id
                    FROM users
                    WHERE id = ?
                    AND LOWER(TRIM(role)) = 'user'
                    LIMIT 1
                `, [
                    oldUserId
                ]);

            if (oldUser.length === 0) {

                throw new Error(
                    "Old user not found."
                );
            }

            const oldEmployeeId =
                String(
                    oldUser[0].employee_id
                ).trim();

            // =================================================
            // NEW USER
            // =================================================

            const [newUser] =
                await connection.query(`
                    SELECT
                        employee_id,
                        username
                    FROM users
                    WHERE id = ?
                    AND LOWER(TRIM(role)) = 'user'
                    AND is_active = TRUE
                    LIMIT 1
                `, [
                    newUserId
                ]);

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

            // =================================================
            // REASSIGN
            // =================================================

            await connection.query(`
                UPDATE claims
                SET
                    assigned_user_id = ?,
                    user_name = ?,
                    updated_at = NOW()
                WHERE
                    TRIM(assigned_user_id)
                    = TRIM(?)
            `, [
                newEmployeeId,
                newUsername,
                oldEmployeeId
            ]);

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
// SAVE USER CLAIMS
// =====================================================

app.post(
    "/save-claims",
    async (req, res) => {

        if (
            !req.session.user ||
            normalizeRole(req.session.user.role) !== "user"
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
            // GET USER CLAIMS
            // =================================================

            const [claims] =
                await connection.query(`
                    SELECT
                        id
                    FROM claims
                    WHERE TRIM(assigned_user_id)
                        = TRIM(?)
                `, [
                    employeeId
                ]);

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
                // NORMALIZE CLAIM TYPE
                // =================================================

                const finalClaimType =
                    normalizeClaimType(
                        claimType
                    );

                // =================================================
                // NORMALIZE STATUS
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
                        String(approveAmountRaw).trim() === ""
                    )
                        ? 0
                        : Number(
                            approveAmountRaw
                        );

                const finalDeductionAmount =
                    (
                        deductionAmountRaw === undefined ||
                        deductionAmountRaw === null ||
                        String(deductionAmountRaw).trim() === ""
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

                await connection.query(`
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
                            = TRIM(?)
                `, [

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
                ]);
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
                <h2>Save Claims Failed</h2>

                <pre>${error.message}</pre>

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
            normalizeRole(req.session.user.role) !== "user"
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
                await db.query(`
                    SELECT

                        c.*,

                        u.id AS employeeid,

                        u.employee_id AS employee_id,

                        u.username AS employee_name,

                        ub.uploaded_at AS upload_time

                    FROM claims c

                    LEFT JOIN users u
                        ON TRIM(c.assigned_user_id)
                        = TRIM(u.employee_id)

                    LEFT JOIN upload_batches ub
                        ON c.upload_batch_id
                        = ub.id

                    WHERE
                        TRIM(c.assigned_user_id)
                        = TRIM(?)

                    ORDER BY
                        c.id DESC
                `, [
                    employeeId
                ]);

            // =================================================
            // USER SUMMARY
            // =================================================

            const [[userSummary]] =
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
                        ) AS total_productivity

                    FROM claims

                    WHERE
                        TRIM(assigned_user_id)
                        = TRIM(?)
                `, [
                    employeeId
                ]);

            // =================================================
            // FORMAT CLAIMS
            // =================================================

            const formattedClaims =
                claims.map(
                    claim => {

                        let formattedDate = "-";

                        let formattedTime = "-";

                        // =================================================
                        // DATE
                        // =================================================

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

                        if (claim.claim_time) {

                            const timeValue =
                                String(
                                    claim.claim_time
                                ).trim();

                            if (
                                /^\d{2}:\d{2}:\d{2}$/
                                    .test(
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
                "USER:",
                employeeId
            );

            console.log(
                "CLAIMS:",
                formattedClaims.length
            );

            console.log(
                "SUMMARY:",
                processSummary
            );

            // =================================================
            // RENDER
            // =================================================

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