CREATE TABLE fact_order_promotion_analysis AS
SELECT
    oi.order_id,
    oi.product_id,
    p.product_name,
    p.product_category,
    p.product_subcategory,
    p.product_category || ' > ' || p.product_subcategory AS subcategory_key,
    p.brand,

    t.order_date,
    CAST(STRFTIME('%Y', t.order_date) AS INTEGER)AS order_year,
    CAST(STRFTIME('%m', t.order_date) AS INTEGER)AS order_month,
    STRFTIME('%Y-%m', t.order_date) AS order_year_month,
    CASE WHEN STRFTIME('%Y', t.order_date) BETWEEN '2022' AND '2025'
         THEN 1 ELSE 0 END AS in_scope_2022_2025,
    t.sales_channel,
    t.marketing_channel,

    t.customer_id,
    t.customer_segment,
    COALESCE(m.age_group, 'Unknown') AS age_group,
    COALESCE(m.gender, 'Unknown') AS gender,
    COALESCE(m.churn_status, 'Unknown') AS churn_status,
 
    oi.quantity,
    oi.unit_price,
    oi.discount_percentage,
    oi.discount_amount,
    oi.gross_sales,
    oi.tax_amount,
    oi.shipping_cost,
    oi.net_sales,
    oi.product_cost,
    oi.profit,
 
    ROUND(oi.gross_sales - oi.discount_amount, 2) AS revenue_clean,
    ROUND(oi.gross_sales - oi.discount_amount - oi.product_cost, 2) AS profit_clean,
    CASE WHEN (oi.gross_sales - oi.discount_amount) > 0
         THEN ROUND((oi.gross_sales - oi.discount_amount - oi.product_cost) * 100.0
                    / (oi.gross_sales - oi.discount_amount), 2)
         ELSE NULL END AS margin_clean_pct,

    CASE WHEN oi.discount_percentage > 0 THEN 1 ELSE 0 END AS has_discount,
    CASE
        WHEN oi.discount_percentage <= 0    THEN '0: 0%'
        WHEN oi.discount_percentage <= 0.10 THEN '1: 0-10%'
        WHEN oi.discount_percentage <= 0.20 THEN '2: 10-20%'
        WHEN oi.discount_percentage <= 0.30 THEN '3: 20-30%'
        ELSE '4: >30%'
    END AS discount_bin
FROM order_items oi
INNER JOIN ecommerce_sales_customer_analytics_150k t
        ON oi.order_id = t.order_id
LEFT  JOIN product_catalog p
        ON oi.product_id = p.product_id
LEFT  JOIN master_customer_summary m
        ON t.customer_id = m.customer_id
WHERE UPPER(TRIM(t.order_status)) = 'COMPLETED' AND CAST(STRFTIME('%Y', t.order_date) AS INTEGER) BETWEEN 2022 AND 2025;