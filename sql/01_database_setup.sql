-- =============================================================================
-- Olist E-Commerce Database Setup & Data Ingestion Script
-- Database: olist_analysis
-- Target MySQL Version: 8.4+ (Apple Silicon / macOS)
-- =============================================================================
-- Instructions for running in MySQL CLI:
-- 1. Ensure local_infile is enabled:
--    SET GLOBAL local_infile = 1;
-- 2. Connect with local-infile enabled:
--    mysql --local-infile=1 -u <user> -p olist_analysis < sql/01_database_setup.sql
-- =============================================================================

USE olist_analysis;

-- Temporarily disable foreign key checks for clean DDL execution & bulk loading
SET FOREIGN_KEY_CHECKS = 0;

-- =============================================================================
-- SECTION 1: DDL - TABLE CREATION
-- Drop existing tables in child-to-parent order to respect FK dependencies
-- =============================================================================

DROP TABLE IF EXISTS reviews;
DROP TABLE IF EXISTS payments;
DROP TABLE IF EXISTS order_items;
DROP TABLE IF EXISTS orders;
DROP TABLE IF EXISTS products;
DROP TABLE IF EXISTS category_translation;
DROP TABLE IF EXISTS sellers;
DROP TABLE IF EXISTS customers;

-- -----------------------------------------------------------------------------
-- Table 1: customers
-- Description: Contains customer details including state, city, and zip prefix.
-- Relationships:
--   - customers.customer_id (PK) <--- (1:N) orders.customer_id (FK)
-- -----------------------------------------------------------------------------
CREATE TABLE customers (
    customer_id              CHAR(32)     NOT NULL,
    customer_unique_id       CHAR(32)     NOT NULL,
    customer_zip_code_prefix VARCHAR(5)  NOT NULL,
    customer_city            VARCHAR(50)  NOT NULL,
    customer_state           CHAR(2)      NOT NULL,
    PRIMARY KEY (customer_id),
    INDEX idx_customers_unique_id (customer_unique_id),
    INDEX idx_customers_zip (customer_zip_code_prefix),
    INDEX idx_customers_state_city (customer_state, customer_city)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -----------------------------------------------------------------------------
-- Table 2: sellers
-- Description: Contains seller location details.
-- Relationships:
--   - sellers.seller_id (PK) <--- (1:N) order_items.seller_id (FK)
-- -----------------------------------------------------------------------------
CREATE TABLE sellers (
    seller_id              CHAR(32)     NOT NULL,
    seller_zip_code_prefix VARCHAR(5)  NOT NULL,
    seller_city            VARCHAR(50)  NOT NULL,
    seller_state           CHAR(2)      NOT NULL,
    PRIMARY KEY (seller_id),
    INDEX idx_sellers_zip (seller_zip_code_prefix),
    INDEX idx_sellers_state_city (seller_state, seller_city)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -----------------------------------------------------------------------------
-- Table 3: category_translation
-- Description: Maps Portuguese product category names to English translations.
-- Relationships:
--   - category_translation.product_category_name (PK) <--- (1:N) products.product_category_name
-- -----------------------------------------------------------------------------
CREATE TABLE category_translation (
    product_category_name         VARCHAR(50) NOT NULL,
    product_category_name_english VARCHAR(50) NOT NULL,
    PRIMARY KEY (product_category_name)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -----------------------------------------------------------------------------
-- Table 4: products
-- Description: Contains product attributes (category, dimensions, weight).
-- Relationships:
--   - products.product_id (PK) <--- (1:N) order_items.product_id (FK)
--   - products.product_category_name ---> category_translation.product_category_name
--     (Note: 610 products have NULL category_name, and 2 categories are unmapped in translation table)
-- -----------------------------------------------------------------------------
CREATE TABLE products (
    product_id                 CHAR(32)    NOT NULL,
    product_category_name      VARCHAR(50) NULL,
    product_name_lenght        INT         NULL,
    product_description_lenght INT         NULL,
    product_photos_qty         INT         NULL,
    product_weight_g           INT         NULL,
    product_length_cm          INT         NULL,
    product_height_cm          INT         NULL,
    product_width_cm           INT         NULL,
    PRIMARY KEY (product_id),
    INDEX idx_products_category (product_category_name)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -----------------------------------------------------------------------------
-- Table 5: orders
-- Description: Core orders table storing purchase timestamps, status, and delivery milestones.
-- Relationships:
--   - orders.customer_id (FK) ---> customers.customer_id (PK)
--   - orders.order_id (PK) <--- (1:N) order_items.order_id (FK)
--   - orders.order_id (PK) <--- (1:N) payments.order_id (FK)
--   - orders.order_id (PK) <--- (1:N) reviews.order_id (FK)
-- -----------------------------------------------------------------------------
CREATE TABLE orders (
    order_id                      CHAR(32)    NOT NULL,
    customer_id                   CHAR(32)    NOT NULL,
    order_status                  VARCHAR(20) NOT NULL,
    order_purchase_timestamp     DATETIME    NOT NULL,
    order_approved_at             DATETIME    NULL,
    order_delivered_carrier_date  DATETIME    NULL,
    order_delivered_customer_date DATETIME    NULL,
    order_estimated_delivery_date DATETIME    NOT NULL,
    PRIMARY KEY (order_id),
    CONSTRAINT fk_orders_customer FOREIGN KEY (customer_id) REFERENCES customers (customer_id),
    INDEX idx_orders_customer (customer_id),
    INDEX idx_orders_status (order_status),
    INDEX idx_orders_purchase_time (order_purchase_timestamp)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -----------------------------------------------------------------------------
-- Table 6: order_items
-- Description: Line items for each order, linking orders to products and sellers.
-- Relationships:
--   - order_items.order_id (FK) ---> orders.order_id (PK)
--   - order_items.product_id (FK) ---> products.product_id (PK)
--   - order_items.seller_id (FK) ---> sellers.seller_id (PK)
-- -----------------------------------------------------------------------------
CREATE TABLE order_items (
    order_id            CHAR(32)       NOT NULL,
    order_item_id       INT            NOT NULL,
    product_id          CHAR(32)       NOT NULL,
    seller_id           CHAR(32)       NOT NULL,
    shipping_limit_date DATETIME       NOT NULL,
    price               DECIMAL(10, 2) NOT NULL,
    freight_value       DECIMAL(10, 2) NOT NULL,
    PRIMARY KEY (order_id, order_item_id),
    CONSTRAINT fk_order_items_order FOREIGN KEY (order_id) REFERENCES orders (order_id),
    CONSTRAINT fk_order_items_product FOREIGN KEY (product_id) REFERENCES products (product_id),
    CONSTRAINT fk_order_items_seller FOREIGN KEY (seller_id) REFERENCES sellers (seller_id),
    INDEX idx_order_items_product (product_id),
    INDEX idx_order_items_seller (seller_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -----------------------------------------------------------------------------
-- Table 7: payments
-- Description: Payment records per order (supports split payments across methods).
-- Relationships:
--   - payments.order_id (FK) ---> orders.order_id (PK)
-- -----------------------------------------------------------------------------
CREATE TABLE payments (
    order_id             CHAR(32)       NOT NULL,
    payment_sequential   INT            NOT NULL,
    payment_type         VARCHAR(20)    NOT NULL,
    payment_installments INT            NOT NULL,
    payment_value        DECIMAL(10, 2) NOT NULL,
    PRIMARY KEY (order_id, payment_sequential),
    CONSTRAINT fk_payments_order FOREIGN KEY (order_id) REFERENCES orders (order_id),
    INDEX idx_payments_type (payment_type)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- -----------------------------------------------------------------------------
-- Table 8: reviews
-- Description: Customer satisfaction scores, review titles, and text comments.
-- Relationships:
--   - reviews.order_id (FK) ---> orders.order_id (PK)
--   - Primary Key is composite (review_id, order_id) because review_id is non-unique
--     (789 review_ids are shared across different orders).
-- -----------------------------------------------------------------------------
CREATE TABLE reviews (
    review_id               CHAR(32) NOT NULL,
    order_id                CHAR(32) NOT NULL,
    review_score            TINYINT  NOT NULL,
    review_comment_title    VARCHAR(100) NULL,
    review_comment_message  TEXT         NULL,
    review_creation_date    DATETIME NOT NULL,
    review_answer_timestamp DATETIME NOT NULL,
    PRIMARY KEY (review_id, order_id),
    CONSTRAINT fk_reviews_order FOREIGN KEY (order_id) REFERENCES orders (order_id),
    INDEX idx_reviews_order (order_id),
    INDEX idx_reviews_score (review_score)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- =============================================================================
-- SECTION 2: DATA INGESTION (LOAD DATA LOCAL INFILE)
-- Adjust relative paths if running MySQL CLI from outside the project root directory.
-- =============================================================================

-- 1. Ingest Customers
LOAD DATA LOCAL INFILE 'olist_customers_dataset.csv'
INTO TABLE customers
FIELDS TERMINATED BY ',' 
ENCLOSED BY '"' 
LINES TERMINATED BY '\n'
IGNORE 1 LINES
(customer_id, customer_unique_id, customer_zip_code_prefix, customer_city, customer_state);

-- 2. Ingest Sellers
LOAD DATA LOCAL INFILE 'olist_sellers_dataset.csv'
INTO TABLE sellers
FIELDS TERMINATED BY ',' 
ENCLOSED BY '"' 
LINES TERMINATED BY '\n'
IGNORE 1 LINES
(seller_id, seller_zip_code_prefix, seller_city, seller_state);

-- 3. Ingest Category Translations
LOAD DATA LOCAL INFILE 'product_category_name_translation.csv'
INTO TABLE category_translation
FIELDS TERMINATED BY ',' 
ENCLOSED BY '"' 
LINES TERMINATED BY '\r\n'
IGNORE 1 LINES
(product_category_name, product_category_name_english);

-- 4. Ingest Products (Converting empty numeric/string fields to NULL)
LOAD DATA LOCAL INFILE 'olist_products_dataset.csv'
INTO TABLE products
FIELDS TERMINATED BY ',' 
ENCLOSED BY '"' 
LINES TERMINATED BY '\n'
IGNORE 1 LINES
(
    product_id,
    @v_category_name,
    @v_name_length,
    @v_desc_length,
    @v_photos_qty,
    @v_weight_g,
    @v_length_cm,
    @v_height_cm,
    @v_width_cm
)
SET 
    product_category_name      = NULLIF(TRIM(@v_category_name), ''),
    product_name_lenght        = NULLIF(TRIM(@v_name_length), ''),
    product_description_lenght = NULLIF(TRIM(@v_desc_length), ''),
    product_photos_qty         = NULLIF(TRIM(@v_photos_qty), ''),
    product_weight_g           = NULLIF(TRIM(@v_weight_g), ''),
    product_length_cm          = NULLIF(TRIM(@v_length_cm), ''),
    product_height_cm          = NULLIF(TRIM(@v_height_cm), ''),
    product_width_cm           = NULLIF(TRIM(@v_width_cm), '');

-- 5. Ingest Orders (Parsing DATETIME strings and handling optional NULL dates)
LOAD DATA LOCAL INFILE 'olist_orders_dataset.csv'
INTO TABLE orders
FIELDS TERMINATED BY ',' 
ENCLOSED BY '"' 
LINES TERMINATED BY '\n'
IGNORE 1 LINES
(
    order_id,
    customer_id,
    order_status,
    @v_purchase_timestamp,
    @v_approved_at,
    @v_delivered_carrier_date,
    @v_delivered_customer_date,
    @v_estimated_delivery_date
)
SET 
    order_purchase_timestamp      = STR_TO_DATE(@v_purchase_timestamp, '%Y-%m-%d %H:%i:%s'),
    order_approved_at             = NULLIF(TRIM(@v_approved_at), ''),
    order_delivered_carrier_date  = NULLIF(TRIM(@v_delivered_carrier_date), ''),
    order_delivered_customer_date = NULLIF(TRIM(@v_delivered_customer_date), ''),
    order_estimated_delivery_date = STR_TO_DATE(@v_estimated_delivery_date, '%Y-%m-%d %H:%i:%s');

-- 6. Ingest Order Items (Parsing shipping limit date)
LOAD DATA LOCAL INFILE 'olist_order_items_dataset.csv'
INTO TABLE order_items
FIELDS TERMINATED BY ',' 
ENCLOSED BY '"' 
LINES TERMINATED BY '\n'
IGNORE 1 LINES
(
    order_id,
    order_item_id,
    product_id,
    seller_id,
    @v_shipping_limit_date,
    price,
    freight_value
)
SET 
    shipping_limit_date = STR_TO_DATE(@v_shipping_limit_date, '%Y-%m-%d %H:%i:%s');

-- 7. Ingest Payments
LOAD DATA LOCAL INFILE 'olist_order_payments_dataset.csv'
INTO TABLE payments
FIELDS TERMINATED BY ',' 
ENCLOSED BY '"' 
LINES TERMINATED BY '\n'
IGNORE 1 LINES
(
    order_id,
    payment_sequential,
    payment_type,
    payment_installments,
    payment_value
);

-- 8. Ingest Reviews (Handling multiline comments, CRLF, and empty titles/messages)
LOAD DATA LOCAL INFILE 'olist_order_reviews_dataset.csv'
INTO TABLE reviews
FIELDS TERMINATED BY ',' 
ENCLOSED BY '"' 
LINES TERMINATED BY '\r\n'
IGNORE 1 LINES
(
    review_id,
    order_id,
    review_score,
    @v_title,
    @v_message,
    @v_creation_date,
    @v_answer_timestamp
)
SET 
    review_comment_title    = NULLIF(TRIM(@v_title), ''),
    review_comment_message  = NULLIF(TRIM(@v_message), ''),
    review_creation_date    = STR_TO_DATE(@v_creation_date, '%Y-%m-%d %H:%i:%s'),
    review_answer_timestamp = STR_TO_DATE(@v_answer_timestamp, '%Y-%m-%d %H:%i:%s');


-- Re-enable foreign key checks after bulk load
SET FOREIGN_KEY_CHECKS = 1;


-- =============================================================================
-- SECTION 3: VALIDATION QUERIES
-- =============================================================================

SHOW TABLES;

SELECT 'customers' AS table_name, COUNT(*) AS row_count FROM customers
UNION ALL
SELECT 'orders', COUNT(*) FROM orders
UNION ALL
SELECT 'order_items', COUNT(*) FROM order_items
UNION ALL
SELECT 'payments', COUNT(*) FROM payments
UNION ALL
SELECT 'reviews', COUNT(*) FROM reviews
UNION ALL
SELECT 'products', COUNT(*) FROM products
UNION ALL
SELECT 'sellers', COUNT(*) FROM sellers
UNION ALL
SELECT 'category_translation', COUNT(*) FROM category_translation;
