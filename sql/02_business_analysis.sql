-- =============================================================================
-- Olist E-Commerce Business Analysis Queries
-- Database: olist_analysis
-- Target MySQL Version: 8.4+ (Apple Silicon / macOS)
-- Author: Data Analyst Portfolio Project
-- =============================================================================
-- Overview:
-- This script contains end-to-end analytical queries covering:
--   1. DATA QUALITY & DATABASE OVERVIEW
--   2. SALES PERFORMANCE
--   3. CUSTOMER ANALYSIS
--   4. PRODUCT & CATEGORY ANALYSIS
--   5. SELLER ANALYSIS
--   6. DELIVERY PERFORMANCE
--   7. CUSTOMER SATISFACTION
--   8. PAYMENT ANALYSIS
-- =============================================================================

USE olist_analysis;

-- =============================================================================
-- SECTION 1: DATA QUALITY & DATABASE OVERVIEW
-- Objectives: Verify table row counts, inspect dataset time horizon, evaluate 
--             order status breakdown, and audit missing data.
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Query 1.1: Database Table Baseline Metrics
-- Question: What are the total record counts across all core entities in the database?
-- -----------------------------------------------------------------------------
SELECT 'customers' AS entity_name, COUNT(*) AS total_records FROM customers
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

-- -----------------------------------------------------------------------------
-- Query 1.2: Order Status Distribution & Completion Rates
-- Question: What is the volume and percentage breakdown of orders across statuses?
-- -----------------------------------------------------------------------------
SELECT 
    order_status,
    COUNT(*) AS total_orders,
    ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS status_percentage
FROM orders
GROUP BY order_status
ORDER BY total_orders DESC;

-- -----------------------------------------------------------------------------
-- Query 1.3: Dataset Horizon & Date Coverage Range
-- Question: What is the overall timeframe (earliest and latest purchase date) in the dataset?
-- -----------------------------------------------------------------------------
SELECT 
    MIN(order_purchase_timestamp) AS earliest_order_date,
    MAX(order_purchase_timestamp) AS latest_order_date,
    TIMESTAMPDIFF(DAY, MIN(order_purchase_timestamp), MAX(order_purchase_timestamp)) AS dataset_span_days
FROM orders;

-- -----------------------------------------------------------------------------
-- Query 1.4: Data Completeness Audit for Delivery & Product Attributes
-- Question: How many orders have missing delivery dates, and how many products lack categories?
-- -----------------------------------------------------------------------------
SELECT 
    COUNT(*) AS total_orders,
    SUM(CASE WHEN order_approved_at IS NULL THEN 1 ELSE 0 END) AS missing_approval_dates,
    SUM(CASE WHEN order_delivered_carrier_date IS NULL THEN 1 ELSE 0 END) AS missing_carrier_dates,
    SUM(CASE WHEN order_delivered_customer_date IS NULL THEN 1 ELSE 0 END) AS missing_customer_delivery_dates,
    SUM(CASE WHEN order_status = 'delivered' AND order_delivered_customer_date IS NULL THEN 1 ELSE 0 END) AS delivered_status_missing_date
FROM orders;


-- =============================================================================
-- SECTION 2: SALES PERFORMANCE
-- Objectives: Track revenue trends over time, calculate Month-over-Month (MoM) 
--             growth rates, evaluate cumulative revenue, and analyze average order values.
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Query 2.1: Monthly Revenue, Order Volume, and Average Order Value (AOV)
-- Question: How do monthly sales revenue, total order volume, and AOV trend over time?
-- Note: Includes delivered orders to measure realized revenue.
-- -----------------------------------------------------------------------------
SELECT 
    DATE_FORMAT(o.order_purchase_timestamp, '%Y-%m') AS year_month,
    COUNT(DISTINCT o.order_id) AS total_orders,
    COUNT(oi.order_item_id) AS total_items_sold,
    ROUND(SUM(oi.price), 2) AS merchandise_revenue,
    ROUND(SUM(oi.freight_value), 2) AS freight_revenue,
    ROUND(SUM(oi.price + oi.freight_value), 2) AS gross_total_revenue,
    ROUND(SUM(oi.price + oi.freight_value) / COUNT(DISTINCT o.order_id), 2) AS average_order_value
FROM orders o
JOIN order_items oi ON o.order_id = oi.order_id
WHERE o.order_status = 'delivered'
GROUP BY DATE_FORMAT(o.order_purchase_timestamp, '%Y-%m')
ORDER BY year_month;

-- -----------------------------------------------------------------------------
-- Query 2.2: Month-over-Month (MoM) Revenue Growth Rate
-- Question: What is the monthly revenue growth percentage compared to the preceding month?
-- -----------------------------------------------------------------------------
WITH monthly_revenue AS (
    SELECT 
        DATE_FORMAT(o.order_purchase_timestamp, '%Y-%m') AS year_month,
        ROUND(SUM(oi.price + oi.freight_value), 2) AS total_revenue
    FROM orders o
    JOIN order_items oi ON o.order_id = oi.order_id
    WHERE o.order_status = 'delivered'
    GROUP BY DATE_FORMAT(o.order_purchase_timestamp, '%Y-%m')
)
SELECT 
    year_month,
    total_revenue,
    LAG(total_revenue, 1) OVER (ORDER BY year_month) AS prior_month_revenue,
    ROUND(
        (total_revenue - LAG(total_revenue, 1) OVER (ORDER BY year_month)) * 100.0 / 
        NULLIF(LAG(total_revenue, 1) OVER (ORDER BY year_month), 0), 2
    ) AS mom_growth_percentage
FROM monthly_revenue
ORDER BY year_month;

-- -----------------------------------------------------------------------------
-- Query 2.3: Cumulative Realized Revenue Over Time
-- Question: What is the cumulative running total of revenue generated over the marketplace lifespan?
-- -----------------------------------------------------------------------------
WITH monthly_sales AS (
    SELECT 
        DATE_FORMAT(o.order_purchase_timestamp, '%Y-%m') AS year_month,
        ROUND(SUM(oi.price + oi.freight_value), 2) AS monthly_revenue
    FROM orders o
    JOIN order_items oi ON o.order_id = oi.order_id
    WHERE o.order_status = 'delivered'
    GROUP BY DATE_FORMAT(o.order_purchase_timestamp, '%Y-%m')
)
SELECT 
    year_month,
    monthly_revenue,
    ROUND(SUM(monthly_revenue) OVER (ORDER BY year_month ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW), 2) AS cumulative_revenue
FROM monthly_sales
ORDER BY year_month;

-- -----------------------------------------------------------------------------
-- Query 2.4: Revenue Loss Impact Analysis from Canceled & Unavailable Orders
-- Question: How much potential revenue was lost due to order cancellations or fulfillment unavailability?
-- -----------------------------------------------------------------------------
SELECT 
    o.order_status,
    COUNT(DISTINCT o.order_id) AS order_count,
    ROUND(SUM(oi.price), 2) AS total_item_price,
    ROUND(SUM(oi.freight_value), 2) AS total_freight,
    ROUND(SUM(oi.price + oi.freight_value), 2) AS total_potential_value
FROM orders o
JOIN order_items oi ON o.order_id = oi.order_id
GROUP BY o.order_status
ORDER BY total_potential_value DESC;


-- =============================================================================
-- SECTION 3: CUSTOMER ANALYSIS
-- Objectives: Evaluate customer geographic concentration, identify repeat vs one-time 
--             buyers, perform RFM segmentation, and track acquisition cohorts.
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Query 3.1a: Top 10 Customer States by Revenue and Customer Volume
-- Question: Which Brazilian states generate the highest revenue and customer volume?
-- -----------------------------------------------------------------------------
SELECT 
    c.customer_state,
    COUNT(DISTINCT c.customer_unique_id) AS unique_customers,
    COUNT(DISTINCT o.order_id) AS total_orders,
    ROUND(SUM(oi.price + oi.freight_value), 2) AS total_spend,
    ROUND(SUM(oi.price + oi.freight_value) / COUNT(DISTINCT c.customer_unique_id), 2) AS spend_per_unique_customer
FROM customers c
JOIN orders o ON c.customer_id = o.customer_id
JOIN order_items oi ON o.order_id = oi.order_id
WHERE o.order_status = 'delivered'
GROUP BY c.customer_state
ORDER BY total_spend DESC
LIMIT 10;

-- -----------------------------------------------------------------------------
-- Query 3.1b: Top 10 Customer Cities by Revenue and Customer Volume
-- Question: Which Brazilian cities generate the highest revenue and customer volume?
-- -----------------------------------------------------------------------------
SELECT 
    c.customer_city,
    c.customer_state,
    COUNT(DISTINCT c.customer_unique_id) AS unique_customers,
    COUNT(DISTINCT o.order_id) AS total_orders,
    ROUND(SUM(oi.price + oi.freight_value), 2) AS total_spend,
    ROUND(SUM(oi.price + oi.freight_value) / COUNT(DISTINCT c.customer_unique_id), 2) AS spend_per_unique_customer
FROM customers c
JOIN orders o ON c.customer_id = o.customer_id
JOIN order_items oi ON o.order_id = oi.order_id
WHERE o.order_status = 'delivered'
GROUP BY c.customer_city, c.customer_state
ORDER BY total_spend DESC
LIMIT 10;

-- -----------------------------------------------------------------------------
-- Query 3.2: Customer Retention - One-Time vs Repeat Purchase Rate
-- Question: What proportion of unique customers place multiple orders vs a single purchase?
-- -----------------------------------------------------------------------------
WITH customer_order_counts AS (
    SELECT 
        c.customer_unique_id,
        COUNT(DISTINCT o.order_id) AS order_count
    FROM customers c
    JOIN orders o ON c.customer_id = o.customer_id
    WHERE o.order_status = 'delivered'
    GROUP BY c.customer_unique_id
)
SELECT 
    CASE 
        WHEN order_count = 1 THEN '1 Order (One-time)'
        WHEN order_count = 2 THEN '2 Orders'
        WHEN order_count = 3 THEN '3 Orders'
        ELSE '4+ Orders (Power Buyers)'
    END AS purchase_frequency_tier,
    COUNT(*) AS customer_count,
    ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS percentage_of_customers
FROM customer_order_counts
GROUP BY 
    CASE 
        WHEN order_count = 1 THEN '1 Order (One-time)'
        WHEN order_count = 2 THEN '2 Orders'
        WHEN order_count = 3 THEN '3 Orders'
        ELSE '4+ Orders (Power Buyers)'
    END
ORDER BY customer_count DESC;

-- -----------------------------------------------------------------------------
-- Query 3.3: RFM Customer Segmentation Model (Recency, Frequency, Monetary)
-- Question: How can customers be segmented into RFM quartiles based on purchase behavior?
-- -----------------------------------------------------------------------------
WITH max_dataset_date AS (
    SELECT MAX(order_purchase_timestamp) AS ref_date FROM orders
),
customer_rfm_raw AS (
    SELECT 
        c.customer_unique_id,
        TIMESTAMPDIFF(DAY, MAX(o.order_purchase_timestamp), m.ref_date) AS recency_days,
        COUNT(DISTINCT o.order_id) AS frequency,
        SUM(oi.price + oi.freight_value) AS monetary_val
    FROM customers c
    JOIN orders o ON c.customer_id = o.customer_id
    JOIN order_items oi ON o.order_id = oi.order_id
    CROSS JOIN max_dataset_date m
    WHERE o.order_status = 'delivered'
    GROUP BY c.customer_unique_id, m.ref_date
)
SELECT 
    customer_unique_id,
    recency_days,
    frequency,
    ROUND(monetary_val, 2) AS monetary_value,
    NTILE(4) OVER (ORDER BY recency_days ASC) AS r_quartile,
    NTILE(4) OVER (ORDER BY monetary_val DESC) AS m_quartile
FROM customer_rfm_raw
ORDER BY monetary_val DESC
LIMIT 20;

-- -----------------------------------------------------------------------------
-- Query 3.4: Monthly Customer Acquisition Cohort Analysis (First Delivered Purchase)
-- Question: How many new unique customers made their first delivered purchase each month?
-- -----------------------------------------------------------------------------
WITH customer_first_purchase AS (
    SELECT 
        c.customer_unique_id,
        MIN(o.order_purchase_timestamp) AS first_order_date
    FROM customers c
    JOIN orders o ON c.customer_id = o.customer_id
    WHERE o.order_status = 'delivered'
    GROUP BY c.customer_unique_id
)
SELECT 
    DATE_FORMAT(first_order_date, '%Y-%m') AS cohort_month,
    COUNT(customer_unique_id) AS new_customers_acquired
FROM customer_first_purchase
GROUP BY DATE_FORMAT(first_order_date, '%Y-%m')
ORDER BY cohort_month;


-- =============================================================================
-- SECTION 4: PRODUCT & CATEGORY ANALYSIS
-- Objectives: Identify top-performing product categories, evaluate individual 
--             best-selling SKUs, and analyze freight-to-price ratios.
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Query 4.1: Top 10 Product Categories by Revenue and Units Sold
-- Question: Which product categories generate the highest revenue and items sold?
-- -----------------------------------------------------------------------------
SELECT 
    COALESCE(t.product_category_name_english, p.product_category_name, 'Uncategorized') AS category_name_english,
    COUNT(oi.order_item_id) AS total_items_sold,
    COUNT(DISTINCT oi.order_id) AS total_orders,
    ROUND(SUM(oi.price), 2) AS total_merchandise_revenue,
    ROUND(AVG(oi.price), 2) AS average_item_price,
    ROUND(AVG(oi.freight_value), 2) AS average_freight_cost
FROM order_items oi
JOIN products p ON oi.product_id = p.product_id
LEFT JOIN category_translation t ON p.product_category_name = t.product_category_name
JOIN orders o ON oi.order_id = o.order_id
WHERE o.order_status = 'delivered'
GROUP BY COALESCE(t.product_category_name_english, p.product_category_name, 'Uncategorized')
ORDER BY total_merchandise_revenue DESC
LIMIT 10;

-- -----------------------------------------------------------------------------
-- Query 4.2: Top 10 Individual Best-Selling Products (SKU Level)
-- Question: What are the top 10 individual products by revenue generation?
-- -----------------------------------------------------------------------------
SELECT 
    p.product_id,
    COALESCE(t.product_category_name_english, p.product_category_name, 'Uncategorized') AS category_name_english,
    COUNT(oi.order_item_id) AS units_sold,
    ROUND(SUM(oi.price), 2) AS total_revenue_generated,
    ROUND(AVG(oi.price), 2) AS unit_price
FROM order_items oi
JOIN products p ON oi.product_id = p.product_id
LEFT JOIN category_translation t ON p.product_category_name = t.product_category_name
JOIN orders o ON oi.order_id = o.order_id
WHERE o.order_status = 'delivered'
GROUP BY p.product_id, COALESCE(t.product_category_name_english, p.product_category_name, 'Uncategorized')
ORDER BY total_revenue_generated DESC
LIMIT 10;

-- -----------------------------------------------------------------------------
-- Query 4.3: Unmapped & Missing Product Categories Audit
-- Question: How much revenue is associated with products lacking an English category mapping?
-- -----------------------------------------------------------------------------
SELECT 
    p.product_category_name AS raw_category_name,
    t.product_category_name_english AS translation_name,
    COUNT(DISTINCT p.product_id) AS distinct_products,
    COUNT(oi.order_item_id) AS items_sold,
    ROUND(SUM(oi.price), 2) AS total_revenue
FROM products p
LEFT JOIN category_translation t ON p.product_category_name = t.product_category_name
JOIN order_items oi ON p.product_id = oi.product_id
WHERE t.product_category_name_english IS NULL
GROUP BY p.product_category_name, t.product_category_name_english
ORDER BY total_revenue DESC;

-- -----------------------------------------------------------------------------
-- Query 4.4: Freight-to-Price Burden Ratio by Category
-- Question: Which product categories carry the highest freight cost relative to product price?
-- -----------------------------------------------------------------------------
SELECT 
    COALESCE(t.product_category_name_english, p.product_category_name, 'Uncategorized') AS category_name_english,
    COUNT(oi.order_item_id) AS total_items,
    ROUND(AVG(oi.price), 2) AS avg_product_price,
    ROUND(AVG(oi.freight_value), 2) AS avg_freight_cost,
    ROUND(AVG(oi.freight_value) * 100.0 / NULLIF(AVG(oi.price), 0), 2) AS freight_to_price_percentage
FROM order_items oi
JOIN products p ON oi.product_id = p.product_id
LEFT JOIN category_translation t ON p.product_category_name = t.product_category_name
JOIN orders o ON oi.order_id = o.order_id
WHERE o.order_status = 'delivered'
GROUP BY COALESCE(t.product_category_name_english, p.product_category_name, 'Uncategorized')
HAVING COUNT(oi.order_item_id) >= 100
ORDER BY freight_to_price_percentage DESC
LIMIT 10;


-- =============================================================================
-- SECTION 5: SELLER ANALYSIS
-- Objectives: Rank top sellers, analyze seller concentration (Pareto principle), 
--             and evaluate seller geographic distribution and shipping latency.
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Query 5.1: Top 10 Sellers by Revenue and Order Volume
-- Question: Who are the top 10 sellers contributing the highest merchandise sales?
-- -----------------------------------------------------------------------------
SELECT 
    s.seller_id,
    s.seller_state,
    s.seller_city,
    COUNT(DISTINCT oi.order_id) AS fulfilled_orders,
    COUNT(oi.order_item_id) AS items_sold,
    ROUND(SUM(oi.price), 2) AS total_merchandise_revenue,
    ROUND(AVG(oi.price), 2) AS avg_item_price
FROM sellers s
JOIN order_items oi ON s.seller_id = oi.seller_id
JOIN orders o ON oi.order_id = o.order_id
WHERE o.order_status = 'delivered'
GROUP BY s.seller_id, s.seller_state, s.seller_city
ORDER BY total_merchandise_revenue DESC
LIMIT 10;

-- -----------------------------------------------------------------------------
-- Query 5.2: Seller Revenue Pareto Analysis & Top 20% Seller Contribution
-- Question: What percentage of total marketplace revenue is generated by the top 20% of sellers (Pareto Principle)?
-- -----------------------------------------------------------------------------
WITH seller_revenue AS (
    SELECT 
        s.seller_id,
        SUM(oi.price) AS total_seller_revenue
    FROM sellers s
    JOIN order_items oi ON s.seller_id = oi.seller_id
    JOIN orders o ON oi.order_id = o.order_id
    WHERE o.order_status = 'delivered'
    GROUP BY s.seller_id
),
seller_ranked AS (
    SELECT 
        seller_id,
        total_seller_revenue,
        ROW_NUMBER() OVER (ORDER BY total_seller_revenue DESC) AS seller_rank,
        COUNT(*) OVER () AS total_seller_count,
        SUM(total_seller_revenue) OVER () AS grand_total_revenue
    FROM seller_revenue
),
seller_pareto AS (
    SELECT 
        seller_id,
        total_seller_revenue,
        grand_total_revenue,
        seller_rank,
        total_seller_count,
        CASE 
            WHEN seller_rank <= CEIL(0.20 * total_seller_count) THEN 'Top 20% Sellers'
            ELSE 'Remaining 80% Sellers'
        END AS seller_tier
    FROM seller_ranked
)
SELECT 
    seller_tier,
    COUNT(seller_id) AS seller_count,
    ROUND(COUNT(seller_id) * 100.0 / MAX(total_seller_count), 2) AS percentage_of_total_sellers,
    ROUND(SUM(total_seller_revenue), 2) AS total_tier_revenue,
    ROUND(SUM(total_seller_revenue) * 100.0 / MAX(grand_total_revenue), 2) AS percentage_of_total_revenue
FROM seller_pareto
GROUP BY seller_tier
ORDER BY total_tier_revenue DESC;

-- -----------------------------------------------------------------------------
-- Query 5.3: Regional Seller vs Customer Imbalance Analysis
-- Question: How are sellers geographically distributed across states compared to buyers?
-- -----------------------------------------------------------------------------
WITH state_sellers AS (
    SELECT seller_state AS state, COUNT(*) AS seller_count FROM sellers GROUP BY seller_state
),
state_customers AS (
    SELECT customer_state AS state, COUNT(DISTINCT customer_unique_id) AS customer_count FROM customers GROUP BY customer_state
)
SELECT 
    COALESCE(sc.state, ss.state) AS state,
    COALESCE(ss.seller_count, 0) AS total_sellers,
    COALESCE(sc.customer_count, 0) AS total_customers,
    ROUND(COALESCE(sc.customer_count, 0) * 1.0 / NULLIF(COALESCE(ss.seller_count, 0), 0), 2) AS customer_to_seller_ratio
FROM state_customers sc
LEFT JOIN state_sellers ss ON sc.state = ss.state
ORDER BY total_customers DESC;


-- =============================================================================
-- SECTION 6: DELIVERY PERFORMANCE
-- Objectives: Calculate order lead times, measure carrier shipping duration, 
--             evaluate on-time vs late delivery rates, and analyze delay gaps by state.
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Query 6.1: Delivery Lead Time Breakdown by Customer State
-- Question: What is the average total delivery lead time (purchase to delivery) across states?
-- -----------------------------------------------------------------------------
SELECT 
    c.customer_state,
    COUNT(o.order_id) AS delivered_orders,
    ROUND(AVG(TIMESTAMPDIFF(DAY, o.order_purchase_timestamp, o.order_delivered_customer_date)), 1) AS avg_total_delivery_days,
    ROUND(AVG(TIMESTAMPDIFF(DAY, o.order_purchase_timestamp, o.order_delivered_carrier_date)), 1) AS avg_fulfillment_days,
    ROUND(AVG(TIMESTAMPDIFF(DAY, o.order_delivered_carrier_date, o.order_delivered_customer_date)), 1) AS avg_carrier_transit_days
FROM orders o
JOIN customers c ON o.customer_id = c.customer_id
WHERE o.order_status = 'delivered'
  AND o.order_delivered_customer_date IS NOT NULL
  AND o.order_delivered_carrier_date IS NOT NULL
GROUP BY c.customer_state
ORDER BY avg_total_delivery_days DESC;

-- -----------------------------------------------------------------------------
-- Query 6.2: On-Time vs Delayed Delivery Rate Performance
-- Question: What percentage of delivered orders arrived on or before the estimated delivery date?
-- -----------------------------------------------------------------------------
SELECT 
    CASE 
        WHEN DATE(order_delivered_customer_date) <= DATE(order_estimated_delivery_date) THEN 'On-Time / Early'
        ELSE 'Late Delivery'
    END AS delivery_performance_status,
    COUNT(*) AS order_count,
    ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS percentage,
    ROUND(AVG(TIMESTAMPDIFF(DAY, order_estimated_delivery_date, order_delivered_customer_date)), 1) AS avg_days_diff_from_estimate
FROM orders
WHERE order_status = 'delivered'
  AND order_delivered_customer_date IS NOT NULL
GROUP BY 
    CASE 
        WHEN DATE(order_delivered_customer_date) <= DATE(order_estimated_delivery_date) THEN 'On-Time / Early'
        ELSE 'Late Delivery'
    END;

-- -----------------------------------------------------------------------------
-- Query 6.3: Monthly Trend of Delivery Estimation Error Gap
-- Question: How accurate were Olist delivery estimations over time?
-- Note: Negative difference means delivery was earlier than estimated.
-- -----------------------------------------------------------------------------
SELECT 
    DATE_FORMAT(order_purchase_timestamp, '%Y-%m') AS purchase_month,
    COUNT(*) AS total_delivered_orders,
    ROUND(AVG(TIMESTAMPDIFF(DAY, order_purchase_timestamp, order_delivered_customer_date)), 1) AS avg_actual_delivery_days,
    ROUND(AVG(TIMESTAMPDIFF(DAY, order_purchase_timestamp, order_estimated_delivery_date)), 1) AS avg_estimated_delivery_days,
    ROUND(AVG(DATEDIFF(order_delivered_customer_date, order_estimated_delivery_date)), 1) AS avg_days_early_or_late
FROM orders
WHERE order_status = 'delivered'
  AND order_delivered_customer_date IS NOT NULL
GROUP BY DATE_FORMAT(order_purchase_timestamp, '%Y-%m')
ORDER BY purchase_month;


-- =============================================================================
-- SECTION 7: CUSTOMER SATISFACTION
-- Objectives: Analyze review rating distribution, measure the impact of delivery 
--             delays on customer satisfaction scores, and rank categories by score.
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Query 7.1: Overall Customer Review Rating Distribution
-- Question: What is the distribution and share of review scores (1 to 5 stars)?
-- -----------------------------------------------------------------------------
SELECT 
    review_score,
    COUNT(*) AS review_count,
    ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS percentage_share
FROM reviews
GROUP BY review_score
ORDER BY review_score DESC;

-- -----------------------------------------------------------------------------
-- Query 7.2: Impact of Delivery Delays on Average Customer Review Score
-- Question: How severely do late deliveries impact average customer review scores?
-- -----------------------------------------------------------------------------
SELECT 
    CASE 
        WHEN DATE(o.order_delivered_customer_date) <= DATE(o.order_estimated_delivery_date) THEN 'On-Time / Early'
        ELSE 'Late Delivery'
    END AS delivery_status,
    COUNT(r.review_id) AS total_reviews,
    ROUND(AVG(r.review_score), 2) AS average_review_score,
    SUM(CASE WHEN r.review_score = 5 THEN 1 ELSE 0 END) AS 5_star_reviews,
    SUM(CASE WHEN r.review_score = 1 THEN 1 ELSE 0 END) AS 1_star_reviews,
    ROUND(SUM(CASE WHEN r.review_score = 1 THEN 1 ELSE 0 END) * 100.0 / COUNT(r.review_id), 2) AS percentage_1_star
FROM orders o
JOIN reviews r ON o.order_id = r.order_id
WHERE o.order_status = 'delivered'
  AND o.order_delivered_customer_date IS NOT NULL
GROUP BY 
    CASE 
        WHEN DATE(o.order_delivered_customer_date) <= DATE(o.order_estimated_delivery_date) THEN 'On-Time / Early'
        ELSE 'Late Delivery'
    END;

-- -----------------------------------------------------------------------------
-- Query 7.3: Top 5 Highest & Lowest Rated Product Categories
-- Question: Which categories achieve the highest and lowest average customer review scores?
-- Note: Deduplicates order-category relationships to prevent review multiplication across multi-item orders.
-- Filtered for categories with at least 100 reviews.
-- -----------------------------------------------------------------------------
WITH order_product_categories AS (
    SELECT DISTINCT 
        oi.order_id,
        COALESCE(t.product_category_name_english, p.product_category_name, 'Uncategorized') AS category_name_english
    FROM order_items oi
    JOIN products p ON oi.product_id = p.product_id
    LEFT JOIN category_translation t ON p.product_category_name = t.product_category_name
),
category_review_scores AS (
    SELECT 
        opc.category_name_english,
        COUNT(r.review_id) AS review_count,
        ROUND(AVG(r.review_score), 2) AS avg_review_score
    FROM order_product_categories opc
    JOIN reviews r ON opc.order_id = r.order_id
    GROUP BY opc.category_name_english
    HAVING COUNT(r.review_id) >= 100
),
ranked_categories AS (
    SELECT 
        category_name_english,
        review_count,
        avg_review_score,
        ROW_NUMBER() OVER (ORDER BY avg_review_score DESC, review_count DESC) AS rank_highest,
        ROW_NUMBER() OVER (ORDER BY avg_review_score ASC, review_count DESC) AS rank_lowest
    FROM category_review_scores
)
SELECT 
    'Highest Rated' AS performance_tier,
    category_name_english,
    review_count,
    avg_review_score
FROM ranked_categories
WHERE rank_highest <= 5

UNION ALL

SELECT 
    'Lowest Rated' AS performance_tier,
    category_name_english,
    review_count,
    avg_review_score
FROM ranked_categories
WHERE rank_lowest <= 5
ORDER BY performance_tier DESC, avg_review_score DESC;


-- -----------------------------------------------------------------------------
-- Query 7.4: Review Comment Text Engagement Rate
-- Question: What percentage of reviews include written comments (title/message) vs rating only?
-- -----------------------------------------------------------------------------
SELECT 
    CASE 
        WHEN review_comment_title IS NOT NULL AND review_comment_message IS NOT NULL THEN 'Both Title & Message'
        WHEN review_comment_message IS NOT NULL THEN 'Message Only'
        WHEN review_comment_title IS NOT NULL THEN 'Title Only'
        ELSE 'Score Only (No Text)'
    END AS comment_engagement_type,
    COUNT(*) AS review_count,
    ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS percentage,
    ROUND(AVG(review_score), 2) AS avg_score_for_engagement_type
FROM reviews
GROUP BY 
    CASE 
        WHEN review_comment_title IS NOT NULL AND review_comment_message IS NOT NULL THEN 'Both Title & Message'
        WHEN review_comment_message IS NOT NULL THEN 'Message Only'
        WHEN review_comment_title IS NOT NULL THEN 'Title Only'
        ELSE 'Score Only (No Text)'
    END
ORDER BY review_count DESC;


-- =============================================================================
-- SECTION 8: PAYMENT ANALYSIS
-- Objectives: Evaluate payment method distribution, analyze installment choices 
--             and their impact on AOV, and detect multi-payment usage.
-- =============================================================================

-- -----------------------------------------------------------------------------
-- Query 8.1: Payment Method Breakdown (Revenue Share & Transaction Volume)
-- Question: What is the relative share of credit card, boleto, voucher, and debit payments?
-- -----------------------------------------------------------------------------
SELECT 
    payment_type,
    COUNT(DISTINCT order_id) AS transaction_count,
    ROUND(SUM(payment_value), 2) AS total_payment_value,
    ROUND(SUM(payment_value) * 100.0 / SUM(SUM(payment_value)) OVER (), 2) AS percentage_of_total_payment,
    ROUND(AVG(payment_value), 2) AS avg_payment_amount
FROM payments
GROUP BY payment_type
ORDER BY total_payment_value DESC;

-- -----------------------------------------------------------------------------
-- Query 8.2: Credit Card Installments Distribution & Order Value Relationship
-- Question: How do credit card installment choices affect average total order payment value?
-- Note: Aggregates payments to the order level first to accurately measure order value.
-- -----------------------------------------------------------------------------
WITH credit_card_orders AS (
    SELECT 
        order_id,
        MAX(payment_installments) AS payment_installments,
        SUM(payment_value) AS total_order_payment_value
    FROM payments
    WHERE payment_type = 'credit_card'
    GROUP BY order_id
)
SELECT 
    payment_installments,
    COUNT(order_id) AS total_orders,
    ROUND(SUM(total_order_payment_value), 2) AS total_payment_value,
    ROUND(AVG(total_order_payment_value), 2) AS avg_order_payment_value
FROM credit_card_orders
GROUP BY payment_installments
ORDER BY payment_installments ASC;

-- -----------------------------------------------------------------------------
-- Query 8.3: Split Payments & Multi-Voucher Usage Detection
-- Question: How many orders use multiple payment methods or split payments across vouchers?
-- -----------------------------------------------------------------------------
WITH order_payment_summary AS (
    SELECT 
        order_id,
        COUNT(payment_sequential) AS payment_splits_count,
        COUNT(DISTINCT payment_type) AS distinct_payment_types,
        SUM(payment_value) AS total_order_payment
    FROM payments
    GROUP BY order_id
)
SELECT 
    CASE 
        WHEN payment_splits_count = 1 THEN 'Single Payment'
        WHEN payment_splits_count > 1 AND distinct_payment_types = 1 THEN 'Split Payment (Same Method/Vouchers)'
        ELSE 'Multi-Method Payment (e.g., Credit + Voucher)'
    END AS payment_behavior,
    COUNT(*) AS order_count,
    ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS percentage_of_orders,
    ROUND(AVG(total_order_payment), 2) AS avg_order_payment_value
FROM order_payment_summary
GROUP BY 
    CASE 
        WHEN payment_splits_count = 1 THEN 'Single Payment'
        WHEN payment_splits_count > 1 AND distinct_payment_types = 1 THEN 'Split Payment (Same Method/Vouchers)'
        ELSE 'Multi-Method Payment (e.g., Credit + Voucher)'
    END
ORDER BY order_count DESC;
