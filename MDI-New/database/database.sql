CREATE DATABASE IF NOT EXISTS claim_management;

USE claim_management;

-- =========================
-- USERS TABLE
-- =========================
CREATE TABLE IF NOT EXISTS users (
    id INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(100) UNIQUE NOT NULL,
    password VARCHAR(255) NOT NULL,
    role ENUM('admin', 'upload', 'user') NOT NULL,
    department VARCHAR(100),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- =========================
-- CLAIMS TABLE
-- =========================
CREATE TABLE IF NOT EXISTS claims (
    id INT AUTO_INCREMENT PRIMARY KEY,

    claim_id VARCHAR(100),
    patient_name VARCHAR(150),
    department VARCHAR(100),

    claim_type VARCHAR(50),
    status VARCHAR(50),

    assigned_user_id INT,

    upload_batch VARCHAR(100),
    uploaded_by INT,

    uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        ON UPDATE CURRENT_TIMESTAMP,

    FOREIGN KEY (assigned_user_id)
        REFERENCES users(id)
        ON DELETE SET NULL,

    FOREIGN KEY (uploaded_by)
        REFERENCES users(id)
        ON DELETE SET NULL
);

-- =========================
-- UPLOAD BATCH TABLE
-- =========================
CREATE TABLE IF NOT EXISTS upload_batches (
    id INT AUTO_INCREMENT PRIMARY KEY,

    file_name VARCHAR(255),
    uploaded_by INT,

    uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

    total_claims INT DEFAULT 0,

    FOREIGN KEY (uploaded_by)
        REFERENCES users(id)
        ON DELETE SET NULL
);

-- =========================
-- DEFAULT ADMIN
-- =========================
INSERT INTO users
(username, password, role, department)
VALUES
('admin', 'admin123', 'admin', 'Management')
ON DUPLICATE KEY UPDATE username = username;

-- =========================
-- DEFAULT UPLOAD LOGIN
-- =========================
INSERT INTO users
(username, password, role, department)
VALUES
('uploader', 'upload123', 'upload', 'Management')
ON DUPLICATE KEY UPDATE username = username;

-- =========================
-- SAMPLE USER
-- =========================
INSERT INTO users
(username, password, role, department)
VALUES
('user1', 'user123', 'user', 'IPD')
ON DUPLICATE KEY UPDATE username = username;