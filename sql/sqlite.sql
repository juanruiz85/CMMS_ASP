-- SQLite CMMS Schema
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS cmms_users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    email TEXT UNIQUE NOT NULL,
    password TEXT NOT NULL,
    role TEXT DEFAULT 'user',
    first_name TEXT,
    last_name TEXT,
    department TEXT,
    avatar TEXT,
    phone TEXT,
    status TEXT CHECK(status IN ('active','inactive')) DEFAULT 'active',
    last_login DATETIME,
    dashboard_config TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS cmms_plants (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT UNIQUE NOT NULL,
    name TEXT NOT NULL,
    description TEXT,
    address TEXT,
    city TEXT,
    state TEXT,
    country TEXT,
    postal_code TEXT,
    manager_id INTEGER,
    phone TEXT,
    email TEXT,
    status TEXT CHECK(status IN ('active','inactive')) DEFAULT 'active',
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (manager_id) REFERENCES cmms_users(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS cmms_assets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT UNIQUE NOT NULL,
    name TEXT NOT NULL,
    description TEXT,
    plant_id INTEGER NOT NULL,
    location TEXT,
    category TEXT,
    manufacturer TEXT,
    model TEXT,
    serial_number TEXT,
    purchase_date DATE,
    installation_date DATE,
    warranty_expiry DATE,
    cost REAL,
    criticality TEXT CHECK(criticality IN ('low','medium','high','critical')) DEFAULT 'medium',
    status TEXT CHECK(status IN ('operational','maintenance','down','retired')) DEFAULT 'operational',
    parent_id INTEGER,
    image_path TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (plant_id) REFERENCES cmms_plants(id) ON DELETE CASCADE,
    FOREIGN KEY (parent_id) REFERENCES cmms_assets(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS cmms_work_orders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT UNIQUE NOT NULL,
    title TEXT NOT NULL,
    description TEXT,
    asset_id INTEGER,
    plant_id INTEGER NOT NULL,
    requester_id INTEGER,
    assigned_to_id INTEGER,
    type TEXT CHECK(type IN ('preventive','corrective','predictive','emergency')) DEFAULT 'corrective',
    priority TEXT CHECK(priority IN ('low','medium','high','urgent')) DEFAULT 'medium',
    status TEXT CHECK(status IN ('open','in_progress','pending','completed','cancelled')) DEFAULT 'open',
    estimated_hours REAL,
    actual_hours REAL,
    estimated_cost REAL,
    actual_cost REAL,
    scheduled_start DATETIME,
    scheduled_end DATETIME,
    actual_start DATETIME,
    actual_end DATETIME,
    completed_at DATETIME,
    closed_by_id INTEGER,
    notes TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (asset_id) REFERENCES cmms_assets(id) ON DELETE SET NULL,
    FOREIGN KEY (plant_id) REFERENCES cmms_plants(id) ON DELETE CASCADE,
    FOREIGN KEY (requester_id) REFERENCES cmms_users(id) ON DELETE SET NULL,
    FOREIGN KEY (assigned_to_id) REFERENCES cmms_users(id) ON DELETE SET NULL,
    FOREIGN KEY (closed_by_id) REFERENCES cmms_users(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS cmms_inventory (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT UNIQUE NOT NULL,
    name TEXT NOT NULL,
    description TEXT,
    category TEXT,
    manufacturer TEXT,
    model TEXT,
    part_number TEXT,
    unit_of_measure TEXT DEFAULT 'unit',
    min_stock REAL DEFAULT 0,
    max_stock REAL DEFAULT 0,
    reorder_point REAL DEFAULT 0,
    unit_cost REAL DEFAULT 0,
    plant_id INTEGER,
    location TEXT,
    bin_location TEXT,
    barcode TEXT,
    image_path TEXT,
    status TEXT CHECK(status IN ('active','inactive')) DEFAULT 'active',
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (plant_id) REFERENCES cmms_plants(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS cmms_inventory_stock (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    inventory_id INTEGER NOT NULL,
    plant_id INTEGER NOT NULL,
    quantity REAL DEFAULT 0,
    reserved_quantity REAL DEFAULT 0,
    available_quantity REAL DEFAULT 0,
    last_updated DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (inventory_id) REFERENCES cmms_inventory(id) ON DELETE CASCADE,
    FOREIGN KEY (plant_id) REFERENCES cmms_plants(id) ON DELETE CASCADE,
    UNIQUE(inventory_id, plant_id)
);

CREATE TABLE IF NOT EXISTS cmms_roles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL,
    description TEXT,
    permissions TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS cmms_settings (
    key_name TEXT PRIMARY KEY,
    value TEXT,
    description TEXT
);

CREATE TABLE IF NOT EXISTS cmms_activity_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    action TEXT NOT NULL,
    description TEXT,
    entity_type TEXT,
    entity_id INTEGER,
    ip_address TEXT,
    user_agent TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (user_id) REFERENCES cmms_users(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS cmms_notifications (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    title TEXT NOT NULL,
    message TEXT,
    type TEXT CHECK(type IN ('info','warning','error','success')) DEFAULT 'info',
    entity_type TEXT,
    entity_id INTEGER,
    is_read INTEGER DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (user_id) REFERENCES cmms_users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS cmms_reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    description TEXT,
    type TEXT CHECK(type IN ('table','chart','dashboard')) DEFAULT 'table',
    query TEXT,
    config TEXT,
    created_by_id INTEGER,
    is_public INTEGER DEFAULT 0,
    is_active INTEGER DEFAULT 1,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (created_by_id) REFERENCES cmms_users(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS cmms_work_order_comments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    work_order_id INTEGER NOT NULL,
    user_id INTEGER,
    comment TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (work_order_id) REFERENCES cmms_work_orders(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id) REFERENCES cmms_users(id)
);

CREATE TABLE IF NOT EXISTS cmms_work_order_time_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    work_order_id INTEGER NOT NULL,
    user_id INTEGER NOT NULL,
    hours REAL NOT NULL,
    description TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (work_order_id) REFERENCES cmms_work_orders(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id) REFERENCES cmms_users(id)
);

CREATE TABLE IF NOT EXISTS cmms_work_order_materials (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    work_order_id INTEGER NOT NULL,
    inventory_id INTEGER NOT NULL,
    quantity REAL NOT NULL,
    cost REAL,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (work_order_id) REFERENCES cmms_work_orders(id) ON DELETE CASCADE,
    FOREIGN KEY (inventory_id) REFERENCES cmms_inventory(id)
);

CREATE TABLE IF NOT EXISTS cmms_inventory_movements (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    inventory_id INTEGER NOT NULL,
    movement_type TEXT NOT NULL,
    quantity REAL NOT NULL,
    previous_quantity REAL,
    new_quantity REAL,
    reason TEXT,
    reference TEXT,
    user_id INTEGER,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (inventory_id) REFERENCES cmms_inventory(id) ON DELETE CASCADE,
    FOREIGN KEY (user_id) REFERENCES cmms_users(id)
);

CREATE TABLE IF NOT EXISTS cmms_purchase_requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    inventory_id INTEGER NOT NULL,
    quantity REAL NOT NULL,
    priority TEXT CHECK(priority IN ('low','medium','high')) DEFAULT 'medium',
    status TEXT CHECK(status IN ('pending','approved','rejected','ordered')) DEFAULT 'pending',
    requested_by_id INTEGER NOT NULL,
    approved_by_id INTEGER,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (inventory_id) REFERENCES cmms_inventory(id),
    FOREIGN KEY (requested_by_id) REFERENCES cmms_users(id),
    FOREIGN KEY (approved_by_id) REFERENCES cmms_users(id)
);