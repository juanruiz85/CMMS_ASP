-- MS SQL Server CMMS Schema
-- Compatible con SQL Server 2016+

IF OBJECT_ID('cmms_users', 'U') IS NULL
CREATE TABLE cmms_users (
    id INT IDENTITY(1,1) PRIMARY KEY,
    username VARCHAR(50) UNIQUE NOT NULL,
    email VARCHAR(100) UNIQUE NOT NULL,
    password VARCHAR(255) NOT NULL,
    role VARCHAR(50) NOT NULL DEFAULT 'user',
    first_name VARCHAR(100),
    last_name VARCHAR(100),
    department VARCHAR(100),
    avatar VARCHAR(255),
    phone VARCHAR(20),
    status VARCHAR(20) DEFAULT 'active',
    last_login DATETIME2,
    dashboard_config NVARCHAR(MAX),
    created_at DATETIME2 DEFAULT SYSUTCDATETIME(),
    updated_at DATETIME2 DEFAULT SYSUTCDATETIME()
);

IF OBJECT_ID('cmms_plants', 'U') IS NULL
CREATE TABLE cmms_plants (
    id INT IDENTITY(1,1) PRIMARY KEY,
    code VARCHAR(50) UNIQUE NOT NULL,
    name VARCHAR(150) NOT NULL,
    description NVARCHAR(MAX),
    address VARCHAR(255),
    city VARCHAR(100),
    state VARCHAR(100),
    country VARCHAR(100),
    postal_code VARCHAR(20),
    manager_id INT,
    phone VARCHAR(50),
    email VARCHAR(100),
    status VARCHAR(20) DEFAULT 'active',
    created_at DATETIME2 DEFAULT SYSUTCDATETIME(),
    updated_at DATETIME2 DEFAULT SYSUTCDATETIME()
);

IF OBJECT_ID('cmms_assets', 'U') IS NULL
CREATE TABLE cmms_assets (
    id INT IDENTITY(1,1) PRIMARY KEY,
    code VARCHAR(100) UNIQUE NOT NULL,
    name VARCHAR(200) NOT NULL,
    description NVARCHAR(MAX),
    plant_id INT NOT NULL,
    location VARCHAR(255),
    category VARCHAR(100),
    manufacturer VARCHAR(150),
    model VARCHAR(100),
    serial_number VARCHAR(100),
    purchase_date DATE,
    installation_date DATE,
    warranty_expiry DATE,
    cost DECIMAL(12,2),
    criticality VARCHAR(20) DEFAULT 'medium',
    status VARCHAR(20) DEFAULT 'operational',
    parent_id INT,
    image_path VARCHAR(255),
    created_at DATETIME2 DEFAULT SYSUTCDATETIME(),
    updated_at DATETIME2 DEFAULT SYSUTCDATETIME()
);

IF OBJECT_ID('cmms_work_orders', 'U') IS NULL
CREATE TABLE cmms_work_orders (
    id INT IDENTITY(1,1) PRIMARY KEY,
    code VARCHAR(50) UNIQUE NOT NULL,
    title VARCHAR(255) NOT NULL,
    description NVARCHAR(MAX),
    asset_id INT,
    plant_id INT NOT NULL,
    requester_id INT,
    assigned_to_id INT,
    type VARCHAR(20) DEFAULT 'corrective',
    priority VARCHAR(20) DEFAULT 'medium',
    status VARCHAR(20) DEFAULT 'open',
    estimated_hours DECIMAL(6,2),
    actual_hours DECIMAL(6,2),
    estimated_cost DECIMAL(12,2),
    actual_cost DECIMAL(12,2),
    scheduled_start DATETIME2,
    scheduled_end DATETIME2,
    actual_start DATETIME2,
    actual_end DATETIME2,
    completed_at DATETIME2,
    closed_by_id INT,
    notes NVARCHAR(MAX),
    created_at DATETIME2 DEFAULT SYSUTCDATETIME(),
    updated_at DATETIME2 DEFAULT SYSUTCDATETIME()
);

IF OBJECT_ID('cmms_inventory', 'U') IS NULL
CREATE TABLE cmms_inventory (
    id INT IDENTITY(1,1) PRIMARY KEY,
    code VARCHAR(100) UNIQUE NOT NULL,
    name VARCHAR(200) NOT NULL,
    description NVARCHAR(MAX),
    category VARCHAR(100),
    manufacturer VARCHAR(150),
    model VARCHAR(100),
    part_number VARCHAR(100),
    unit_of_measure VARCHAR(20) DEFAULT 'unit',
    min_stock DECIMAL(10,2) DEFAULT 0,
    max_stock DECIMAL(10,2) DEFAULT 0,
    reorder_point DECIMAL(10,2) DEFAULT 0,
    unit_cost DECIMAL(12,2) DEFAULT 0,
    plant_id INT,
    location VARCHAR(100),
    bin_location VARCHAR(50),
    barcode VARCHAR(100),
    image_path VARCHAR(255),
    status VARCHAR(20) DEFAULT 'active',
    created_at DATETIME2 DEFAULT SYSUTCDATETIME(),
    updated_at DATETIME2 DEFAULT SYSUTCDATETIME()
);

IF OBJECT_ID('cmms_inventory_stock', 'U') IS NULL
CREATE TABLE cmms_inventory_stock (
    id INT IDENTITY(1,1) PRIMARY KEY,
    inventory_id INT NOT NULL,
    plant_id INT NOT NULL,
    quantity DECIMAL(10,2) DEFAULT 0,
    reserved_quantity DECIMAL(10,2) DEFAULT 0,
    available_quantity DECIMAL(10,2) DEFAULT 0,
    last_updated DATETIME2 DEFAULT SYSUTCDATETIME(),
    UNIQUE(inventory_id, plant_id)
);

IF OBJECT_ID('cmms_roles', 'U') IS NULL
CREATE TABLE cmms_roles (
    id INT IDENTITY(1,1) PRIMARY KEY,
    name VARCHAR(100) UNIQUE NOT NULL,
    description NVARCHAR(255),
    permissions NVARCHAR(MAX),
    created_at DATETIME2 DEFAULT SYSUTCDATETIME()
);

IF OBJECT_ID('cmms_settings', 'U') IS NULL
CREATE TABLE cmms_settings (
    [key] VARCHAR(150) PRIMARY KEY,
    [value] NVARCHAR(MAX),
    description NVARCHAR(255),
    updated_at DATETIME2 DEFAULT SYSUTCDATETIME()
);

IF OBJECT_ID('cmms_activity_logs', 'U') IS NULL
CREATE TABLE cmms_activity_logs (
    id INT IDENTITY(1,1) PRIMARY KEY,
    user_id INT,
    action VARCHAR(255) NOT NULL,
    description NVARCHAR(MAX),
    entity_type VARCHAR(100),
    entity_id INT,
    ip_address VARCHAR(45),
    user_agent NVARCHAR(255),
    created_at DATETIME2 DEFAULT SYSUTCDATETIME()
);

IF OBJECT_ID('cmms_notifications', 'U') IS NULL
CREATE TABLE cmms_notifications (
    id INT IDENTITY(1,1) PRIMARY KEY,
    user_id INT NOT NULL,
    title VARCHAR(255) NOT NULL,
    message NVARCHAR(MAX),
    type VARCHAR(20) DEFAULT 'info',
    entity_type VARCHAR(100),
    entity_id INT,
    is_read BIT DEFAULT 0,
    created_at DATETIME2 DEFAULT SYSUTCDATETIME(),
    FOREIGN KEY (user_id) REFERENCES cmms_users(id) ON DELETE CASCADE
);

IF OBJECT_ID('cmms_reports', 'U') IS NULL
CREATE TABLE cmms_reports (
    id INT IDENTITY(1,1) PRIMARY KEY,
    name VARCHAR(200) NOT NULL,
    description NVARCHAR(MAX),
    type VARCHAR(20) DEFAULT 'table',
    query NVARCHAR(MAX),
    config NVARCHAR(MAX),
    created_by_id INT,
    is_public BIT DEFAULT 0,
    is_active BIT DEFAULT 1,
    created_at DATETIME2 DEFAULT SYSUTCDATETIME(),
    updated_at DATETIME2 DEFAULT SYSUTCDATETIME()
);

IF OBJECT_ID('cmms_work_order_comments', 'U') IS NULL
CREATE TABLE cmms_work_order_comments (
    id INT IDENTITY(1,1) PRIMARY KEY,
    work_order_id INT NOT NULL,
    user_id INT,
    comment NVARCHAR(MAX),
    created_at DATETIME2 DEFAULT SYSUTCDATETIME(),
    FOREIGN KEY (work_order_id) REFERENCES cmms_work_orders(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id) REFERENCES cmms_users(id)
);

IF OBJECT_ID('cmms_work_order_time_logs', 'U') IS NULL
CREATE TABLE cmms_work_order_time_logs (
    id INT IDENTITY(1,1) PRIMARY KEY,
    work_order_id INT NOT NULL,
    user_id INT NOT NULL,
    hours DECIMAL(6,2) NOT NULL,
    description NVARCHAR(MAX),
    created_at DATETIME2 DEFAULT SYSUTCDATETIME(),
    FOREIGN KEY (work_order_id) REFERENCES cmms_work_orders(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id) REFERENCES cmms_users(id)
);

IF OBJECT_ID('cmms_work_order_materials', 'U') IS NULL
CREATE TABLE cmms_work_order_materials (
    id INT IDENTITY(1,1) PRIMARY KEY,
    work_order_id INT NOT NULL,
    inventory_id INT NOT NULL,
    quantity DECIMAL(10,2) NOT NULL,
    cost DECIMAL(12,2),
    created_at DATETIME2 DEFAULT SYSUTCDATETIME(),
    FOREIGN KEY (work_order_id) REFERENCES cmms_work_orders(id) ON DELETE CASCADE,
    FOREIGN KEY (inventory_id) REFERENCES cmms_inventory(id)
);

IF OBJECT_ID('cmms_inventory_movements', 'U') IS NULL
CREATE TABLE cmms_inventory_movements (
    id INT IDENTITY(1,1) PRIMARY KEY,
    inventory_id INT NOT NULL,
    movement_type VARCHAR(50) NOT NULL,
    quantity DECIMAL(10,2) NOT NULL,
    previous_quantity DECIMAL(10,2),
    new_quantity DECIMAL(10,2),
    reason NVARCHAR(MAX),
    reference VARCHAR(100),
    user_id INT,
    created_at DATETIME2 DEFAULT SYSUTCDATETIME(),
    FOREIGN KEY (inventory_id) REFERENCES cmms_inventory(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id) REFERENCES cmms_users(id)
);

IF OBJECT_ID('cmms_purchase_requests', 'U') IS NULL
CREATE TABLE cmms_purchase_requests (
    id INT IDENTITY(1,1) PRIMARY KEY,
    inventory_id INT NOT NULL,
    quantity DECIMAL(10,2) NOT NULL,
    priority VARCHAR(20) DEFAULT 'medium',
    status VARCHAR(20) DEFAULT 'pending',
    requested_by_id INT NOT NULL,
    approved_by_id INT,
    created_at DATETIME2 DEFAULT SYSUTCDATETIME(),
    updated_at DATETIME2 DEFAULT SYSUTCDATETIME(),
    FOREIGN KEY (inventory_id) REFERENCES cmms_inventory(id),
    FOREIGN KEY (requested_by_id) REFERENCES cmms_users(id),
    FOREIGN KEY (approved_by_id) REFERENCES cmms_users(id)
);